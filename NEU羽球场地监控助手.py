# -*- coding: utf-8 -*-  # 文件编码为 UTF-8
# NEU 场地监控脚本 - GUI 版，配置自动保存，实时日志输出到界面和文件
import os
import time
import threading
import logging
import json
import random
from datetime import datetime
import tkinter as tk
from tkinter import ttk
import smtplib
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

CONFIG_FILE = 'config.json'
DEFAULT_SLOTS = ['08:00-09:00','09:00-10:00','10:00-11:00','11:00-12:00',
                 '12:00-14:00','14:00-16:00','14:00-15:30','15:30-17:00',
                 '16:00-17:00','16:00-18:00','17:00-18:00','18:00-19:00',
                 '18:00-20:00','19:00-20:00','20:00-21:00']

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    def emit(self, record):
        msg = self.format(record)
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.configure(state='disabled')
        self.text_widget.see(tk.END)

# 日志初始化
def setup_logging(text_widget=None):
    os.makedirs('logs', exist_ok=True)
    start_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    path = os.path.join('logs', f'{start_time}.log')
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    fh = logging.FileHandler(path, encoding='utf-8')
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    if text_widget:
        th = TextHandler(text_widget)
        th.setFormatter(fmt)
        logger.addHandler(th)
    logging.info(f'日志输出到 {path}')

# 邮件发送，不记录密码
def send_email(sub, body, server, port, user, pwd, to):
    logging.info('发送邮件中...')
    msg = MIMEText(body, 'html', 'utf-8')
    msg['Subject'], msg['From'], msg['To'] = sub, user, to
    try:
        with smtplib.SMTP(server, port) as s:
            s.starttls()
            s.login(user, pwd)
            s.send_message(msg)
        logging.info(f'邮件已发送: {sub}')
    except Exception as e:
        logging.error(f'发送邮件失败: {e}')

# 浏览器初始化
def init_driver(debug):
    logging.info('初始化浏览器')
    opt = ChromeOptions()
    opt.add_argument('--disable-blink-features=AutomationControlled')
    if not debug:
        opt.add_argument('--headless')
    d = webdriver.Chrome(options=opt)
    d.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
        'source': "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
    })
    return d

# 登录并打开面板
def login_and_open_panel(d, url, user, pwd):
    logging.info('执行登录')
    d.get(url)
    w = WebDriverWait(d, 20)
    w.until(EC.visibility_of_element_located((By.ID, 'un'))).send_keys(user)
    d.find_element(By.ID, 'pd').send_keys(pwd)
    w.until(EC.element_to_be_clickable((By.ID, 'index_login_btn'))).click()
    w.until(EC.element_to_be_clickable((By.CLASS_NAME, 'reserve_button'))).click()
    w.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class,'selectList') and contains(@class,'sectionNotes')]")))
    logging.info('面板加载完毕')

# 监测并发送邮件，延迟随机化
def monitor_slots(d, courts, slots, base_interval, count, mail_cfg):
    for attempt in range(1, count+1):
        logging.info(f'第{attempt}次检查')
        pans = d.find_elements(By.XPATH, "//div[contains(@class,'selectList') and contains(@class,'sectionNotes')]")
        ok = []
        for i in courts:
            if i-1 < len(pans):
                for el in pans[i-1].find_elements(By.XPATH, ".//div[contains(@class,'TimeDiv')]//li"):
                    if any(s in el.text and '可用' in el.text for s in slots):
                        ok.append((i, el.text))
                        break
        if ok:
            body = '<h3>可用：</h3>' + ''.join(f'<p>场地{i}:{t}</p>' for i, t in ok)
            send_email('NEU通知', body, *mail_cfg)
            logging.info('发现可用并已发送邮件，终止监测')
            return
        delay = base_interval * random.uniform(0.8, 1.2)
        logging.info(f'无可用，延迟{delay:.2f}s后重试')
        time.sleep(delay)
        d.refresh()
    logging.info('刷新次数已用完，结束监测')

# 配置读写
def load_config():
    try:
        return json.load(open(CONFIG_FILE, 'r', encoding='utf-8'))
    except:
        return {}

def save_config(cfg):
    json.dump(cfg, open(CONFIG_FILE, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('NEU场地监控')
        self.cfg = load_config()
        self.build_ui()

    def build_ui(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 状态信息区提前定义方便后续使用
        status_fr = ttk.LabelFrame(main_frame, text='状态信息', padding=10)
        status_fr.grid(row=99, column=0, columnspan=2, pady=(10,0), sticky='nsew')
        self.status_text = tk.Text(status_fr, height=8, wrap=tk.WORD, font=('微软雅黑',9), state='disabled')
        self.status_text.pack(fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(status_fr, command=self.status_text.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=sb.set)

        setup_logging(self.status_text)

        def log_change(key, val):
            if '密码' not in key:
                logging.info(f'{key} 设置为 {val}')

        # 登录配置
        ttk.Label(main_frame, text='登录 URL').grid(row=0, column=0)
        ttk.Label(main_frame, text='http://book.neu.edu.cn/booking/page/selectPeList').grid(row=0, column=1)
        self.url = 'http://book.neu.edu.cn/booking/page/selectPeList'; self.entries = {}
        for idx, key in enumerate(['用户名','登录密码'], 1):
            ttk.Label(main_frame, text=key).grid(row=idx, column=0)
            e = ttk.Entry(main_frame, show='*' if '密码' in key else None)
            e.insert(0, self.cfg.get(key,'')); e.grid(row=idx, column=1)
            e.bind('<FocusOut>', lambda ev, k=key, e=e: log_change(k, e.get()))
            self.entries[key] = e

        # 场地选择
        ttk.Label(main_frame, text='场地').grid(row=3, column=0)
        self.courts = {}
        cf = ttk.Frame(main_frame); cf.grid(row=3, column=1)
        for i in range(1,13):
            v = tk.BooleanVar(value=self.cfg.get(f'场地{i}', True))
            cb = ttk.Checkbutton(cf, text=str(i), variable=v)
            cb.grid(row=(i-1)//6, column=(i-1)%6)
            v.trace_add('write', lambda *a, i=i, v=v: log_change(f'场地{i}', v.get()))
            self.courts[i] = v

        # 时段选择
        ttk.Label(main_frame, text='时段').grid(row=4, column=0)
        self.slots = {}
        sf = ttk.Frame(main_frame); sf.grid(row=4, column=1)
        for idx, s in enumerate(DEFAULT_SLOTS):
            v = tk.BooleanVar(value=self.cfg.get(f'时段:{s}', idx<3))
            cb = ttk.Checkbutton(sf, text=s, variable=v)
            cb.grid(row=idx//5, column=idx%5)
            v.trace_add('write', lambda *a, s=s, v=v: log_change(f'时段:{s}', v.get()))
            self.slots[s] = v

        # 刷新间隔与重试次数
        row=5
        for key, defv in [('刷新间隔(s)', '5'), ('最大重试次数','10')]:
            ttk.Label(main_frame, text=key).grid(row=row, column=0)
            e = ttk.Entry(main_frame)
            e.insert(0, self.cfg.get(key, defv)); e.grid(row=row, column=1)
            e.bind('<FocusOut>', lambda ev, k=key, e=e: log_change(k, e.get()))
            self.entries[key] = e; row+=1

        # 邮件配置
        for key, defv in [('SMTP服务器',''),('端口','587'),('邮箱',''),('SMTP密码',''),('收件','')]:
            ttk.Label(main_frame, text=key).grid(row=row, column=0)
            e = ttk.Entry(main_frame, show='*' if '密码' in key else None)
            e.insert(0, self.cfg.get(key, defv)); e.grid(row=row, column=1)
            e.bind('<FocusOut>', lambda ev, k=key, e=e: log_change(k, e.get()))
            self.entries[key] = e; row+=1

        # 调试模式
        self.debug = tk.BooleanVar(value=self.cfg.get('调试模式', False))
        ttk.Checkbutton(main_frame, text='调试模式（打开浏览器可视化）', variable=self.debug).grid(row=row, column=0, columnspan=2)
        self.debug.trace_add('write', lambda *a: log_change('调试模式', self.debug.get()))
        row += 1

        # 底部版权
        ttk.Label(main_frame, text='© 2025 NEU 监控助手', font=('微软雅黑',8)).grid(row=row+1, column=0, columnspan=2, pady=(5,0))
        main_frame.rowconfigure(row, weight=1)

        # 启动按钮
        btn = ttk.Button(main_frame, text='启动', command=self.start)
        btn.grid(row=row+2, column=0, columnspan=2, pady=(10,0))

    def start(self):
        cfg = {k: e.get() for k, e in self.entries.items()}
        for i,v in self.courts.items(): cfg[f'场地{i}'] = v.get()
        for s,v in self.slots.items(): cfg[f'时段:{s}'] = v.get()
        cfg['调试模式'] = self.debug.get()
        save_config(cfg)
        logging.info('开始监控任务')
        driver = init_driver(self.debug.get())
        login_and_open_panel(driver, self.url, cfg['用户名'], cfg['登录密码'])
        mail_cfg = [cfg['SMTP服务器'], int(cfg['端口']), cfg['邮箱'], cfg['SMTP密码'], cfg['收件']]
        base_interval = float(cfg['刷新间隔(s)'])
        count = int(cfg['最大重试次数'])
        threading.Thread(target=monitor_slots, args=(driver,
            [i for i,v in self.courts.items() if v.get()],
            [s for s,v in self.slots.items() if v.get()],
            base_interval, count, mail_cfg), daemon=True).start()

if __name__ == '__main__':
    App().mainloop()
