# -*- coding: utf-8 -*-  # 文件编码为 UTF-8
# NEU 场地监控脚本 - GUI 版，配置自动保存，实时日志输出到界面和文件
import os
import sys
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
DEFAULT_SLOTS = [
    '08:00-09:00','09:00-10:00','10:00-11:00','11:00-12:00',
    '12:00-14:00','14:00-16:00','14:00-15:30','15:30-17:00',
    '16:00-17:00','16:00-18:00','17:00-18:00','18:00-19:00',
    '18:00-20:00','19:00-20:00','20:00-21:00'
]
# 重新登录间隔：3小时（毫秒）
RELOGIN_INTERVAL = 3 * 60 * 60 * 1000

# 日志处理，将日志写入 Text
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

# 发送邮件（增强异常处理）
def send_email(sub, body, server, port, user, pwd, to):
    logging.info('发送邮件中...')
    msg = MIMEText(body, 'html', 'utf-8')
    msg['Subject'], msg['From'], msg['To'] = sub, user, to
    try:
        with smtplib.SMTP(server, port, timeout=10) as s:
            s.starttls()
            s.login(user, pwd)
            s.send_message(msg)
        logging.info(f'邮件已发送: {sub}')
    except smtplib.SMTPResponseException as e:
        # 部分服务器在发送后断开
        if e.smtp_code < 0:
            logging.warning(f'SMTP 连接断开，邮件可能已发送: {e.smtp_code} - {e.smtp_error}')
        else:
            logging.error(f'SMTP 响应错误: {e.smtp_code} - {e.smtp_error}')
    except smtplib.SMTPException as e:
        logging.error(f'SMTP 错误: {e}')
    except Exception as e:
        logging.error(f'发送邮件失败（未知异常）: {e}')

# 浏览器初始化
def init_driver(debug):
    logging.info('初始化浏览器')
    opt = ChromeOptions()
    opt.add_argument('--disable-blink-features=AutomationControlled')
    if not debug:
        opt.add_argument('--headless')
    d = webdriver.Chrome(options=opt)
    # 隐藏自动化痕迹
    try:
        d.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': "Object.defineProperty(navigator,'webdriver',{get:() => undefined})"
        })
    except Exception:
        # 有些 ChromeDriver 版本/环境可能不支持 execute_cdp_cmd
        pass
    return d

# 登录并打开监控面板
def login_and_open_panel(d, url, user, pwd, verification_code=None):
    logging.info('执行登录')
    d.get(url)
    try:
        # 输入用户名和密码
        w = WebDriverWait(d, 10)
        w.until(EC.visibility_of_element_located((By.ID, 'un'))).send_keys(user)
        d.find_element(By.ID, 'pd').send_keys(pwd)

        # 点击登录按钮
        d.find_element(By.ID, 'index_login_btn').click()

        # 输入验证码（如果提供）
        if verification_code:
            verification_input = WebDriverWait(d, 10).until(EC.visibility_of_element_located((By.ID, 'PM1')))
            verification_input.send_keys(verification_code)
            # 点击登录按钮
            d.find_element(By.ID, 'index_login_btn').click()

        # 等待并进入监控面板
        w.until(EC.element_to_be_clickable((By.CLASS_NAME, 'reserve_button'))).click()
        w.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class,'selectList') and contains(@class,'sectionNotes')]")))
        logging.info('面板加载完毕')
    except Exception:
        logging.error('用户名或密码错误，或者页面未按预期加载，无法访问目标页面')
        raise

# 持续监测并发送通知（改为使用 driver_getter + stop_event，使得可以安全重启浏览器）
def monitor_slots(driver_getter, courts, slots, base_interval, max_retry, mail_cfg, stop_event):
    retry = 0
    prev_state = None  # None 表示首次检查
    while not stop_event.is_set():
        retry += 1
        logging.info(f'第{retry}次检查')
        d = driver_getter()
        if d is None:
            logging.info('浏览器未准备好，等待 1s')
            time.sleep(1)
            continue

        try:
            pans = d.find_elements(By.XPATH, "//div[contains(@class,'selectList') and contains(@class,'sectionNotes')]")
        except Exception as e:
            logging.warning(f'获取页面元素失败（可能是浏览器已重启或连接断开）: {e}')
            # 等待短时间，进入下一循环以便重试或等待 restart 完成
            time.sleep(2)
            continue

        # 构建当前状态：dict {场地号: [可用时段文本, ...]}
        curr_state = {}
        for i in courts:
            available_list = []
            if i-1 < len(pans):
                lis = pans[i-1].find_elements(By.XPATH, ".//div[contains(@class,'TimeDiv')]//li")
                for el in lis:
                    text = el.text.strip()
                    if any(s in text and '可用' in text for s in slots):
                        available_list.append(text)
            curr_state[i] = available_list

        # 构建全局当前可用列表（底部显示一次）：格式为 "场地X: 时段文本"
        overall_current = []
        for i, cur in curr_state.items():
            for c in cur:
                overall_current.append(f'场地{i}: {c}')

        # 计算变化：首次检查（prev_state is None）强制通知；否则比对每个场地新增/取消
        notify = False
        changes = []  # 存放 (场地号, added_set, removed_set)
        if prev_state is None:
            # 修改：首次检查如果全站点无可用则不发送通知（避免每次启动时收到“无变化/无可用”邮件）
            if overall_current:
                notify = True
                for i, cur in curr_state.items():
                    added = set(cur)
                    removed = set()
                    if added:
                        changes.append((i, added, removed))
            else:
                logging.info('首次检查：无可用时段，跳过首次通知')
        else:
            for i, cur in curr_state.items():
                prev_set = set(prev_state.get(i, []))
                curr_set = set(cur)
                added = curr_set - prev_set
                removed = prev_set - curr_set
                if added or removed:
                    notify = True
                    changes.append((i, added, removed))

        # 发送邮件（若需要）
        if notify:
            subject = f'NEU场地状态更新 - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
            if not changes:
                body = '<html><body><h3>场地状态检查（无变化/无可用）</h3></body></html>'
            else:
                body = '<html><body>'
                body += '<h3>场地变更详情（上：每个场地的新增/取消，下：当前全部可用总览）</h3>'
                # 列出每个发生变化的场地（左：新增；同时显示取消）
                for i, added, removed in changes:
                    body += f'<div><strong>场地 {i}</strong></div>'
                    # 新增
                    if added:
                        body += '<div>新增：<ul>'
                        for a in sorted(added):
                            body += f'<li>{a}</li>'
                        body += '</ul></div>'
                    else:
                        body += '<div>新增：—</div>'
                    # 取消（若有）
                    if removed:
                        body += '<div>取消：<ul>'
                        for r in sorted(removed):
                            body += f'<li>{r}</li>'
                        body += '</ul></div>'
                # 分隔并输出一次性全站点总览（右列现在只输出一次在这里）
                body += '<hr/>'
                body += '<h3>当前全部可用（全站点总览）</h3>'
                if overall_current:
                    body += '<ul>'
                    for oc in sorted(overall_current):
                        body += f'<li>{oc}</li>'
                    body += '</ul>'
                else:
                    body += '<div>无</div>'
                body += '</body></html>'

            try:
                send_email(subject, body, *mail_cfg)
                logging.info('检测到变化，已发送通知')
            except Exception as e:
                logging.error(f'发送通知失败: {e}')
        else:
            logging.info('本轮未检测到场地可用时段变化，无需通知')

        # 更新前一状态
        prev_state = curr_state

        # 随机延迟，防止固定频率被识别
        delay = base_interval * random.uniform(0.8, 1.2)
        logging.info(f'延迟{delay:.2f}s后继续监测')

        # 在等待过程中也要响应 stop_event
        slept = 0.0
        while slept < delay:
            if stop_event.is_set():
                logging.info('检测线程收到停止信号，退出循环')
                return
            time.sleep(min(1.0, delay - slept))
            slept += min(1.0, delay - slept)

        try:
            d.refresh()
        except Exception as e:
            logging.warning(f'刷新页面失败: {e}')

        if retry >= max_retry:
            retry = 0
            logging.info('达到最大重试次数，继续循环监控')

# 配置读写
def load_config():
    try:
        return json.load(open(CONFIG_FILE, 'r', encoding='utf-8'))
    except:
        return {}

def save_config(cfg):
    json.dump(cfg, open(CONFIG_FILE, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)

# 主应用界面
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('NEU场地监控')
        self.cfg = load_config()
        self.driver = None
        self._stop_event = threading.Event()  # 用于控制监控线程停止/重启
        self.monitor_thread = None
        self.monitor_params = None  # 存放当前监控线程使用的参数，以便重启时复用
        self.build_ui()
        self.protocol('WM_DELETE_WINDOW', self.on_close)

    def save_and_log_change(self, key, value):
        # 不在日志中输出密码明文
        display_value = value
        if isinstance(key, str) and ('密码' in key or 'password' in key.lower()):
            display_value = '***'
        logging.info(f'配置变更: {key} = {display_value}')
        # 更新内存配置并持久化
        self.cfg[key] = value
        try:
            save_config(self.cfg)
        except Exception as e:
            logging.error(f'保存配置失败: {e}')

    def build_ui(self):
        main = ttk.Frame(self)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

        self.config_widgets = []
        # 登录配置
        ttk.Label(main, text='登录 URL').grid(row=0, column=0, sticky='e')
        ttk.Label(main, text='http://book.neu.edu.cn/booking/page/selectPeList').grid(row=0, column=1, sticky='w')
        self.url = 'http://book.neu.edu.cn/booking/page/selectPeList'
        self.entries = {}
        # 使用 StringVar 来便于追踪变化
        for idx, key in enumerate(['用户名','登录密码'], start=1):
            ttk.Label(main, text=key).grid(row=idx, column=0, sticky='e')
            var = tk.StringVar(value=self.cfg.get(key,''))
            e = ttk.Entry(main, textvariable=var, show='*' if '密码' in key else None)
            e.grid(row=idx, column=1, sticky='we')
            self.entries[key] = var
            self.config_widgets.append(e)
            # 追踪修改并记录日志（自动保存）
            var.trace_add('write', lambda *args, k=key, v=var: self.save_and_log_change(k, v.get()))

        # 场地选择
        ttk.Label(main, text='场地').grid(row=3, column=0, sticky='ne')
        cf = ttk.Frame(main); cf.grid(row=3, column=1, sticky='w')
        self.courts = {}
        for i in range(1,13):
            v = tk.BooleanVar(value=self.cfg.get(f'场地:{i}', True))
            cb = ttk.Checkbutton(cf, text=str(i), variable=v)
            cb.grid(row=(i-1)//6, column=(i-1)%6, sticky='w')
            self.courts[i] = v
            self.config_widgets.append(cb)
            # 追踪并记录日志与持久化
            v.trace_add('write', lambda *args, idx=i, var=v: self.save_and_log_change(f'场地:{idx}', bool(var.get())))

        # 时段选择
        ttk.Label(main, text='时段').grid(row=4, column=0, sticky='ne')
        sf = ttk.Frame(main); sf.grid(row=4, column=1, sticky='w')
        self.slots = {}
        for j, s in enumerate(DEFAULT_SLOTS):
            v = tk.BooleanVar(value=self.cfg.get(f'时段:{s}', j<3))
            cb = ttk.Checkbutton(sf, text=s, variable=v)
            cb.grid(row=j//5, column=j%5, sticky='w')
            self.slots[s] = v
            self.config_widgets.append(cb)
            v.trace_add('write', lambda *args, ss=s, var=v: self.save_and_log_change(f'时段:{ss}', bool(var.get())))

        # 参数输入
        row = 5
        for key, dv in [('刷新间隔(s)','5'),('最大重试次数','10')]:
            ttk.Label(main, text=key).grid(row=row, column=0, sticky='e')
            var = tk.StringVar(value=self.cfg.get(key, dv))
            e = ttk.Entry(main, textvariable=var); e.grid(row=row, column=1, sticky='we')
            self.entries[key] = var
            self.config_widgets.append(e)
            var.trace_add('write', lambda *args, k=key, v=var: self.save_and_log_change(k, v.get()))
            row += 1
        for key, dv in [('SMTP服务器',''),('端口','587'),('邮箱',''),('SMTP密码',''),('收件','')]:
            ttk.Label(main, text=key).grid(row=row, column=0, sticky='e')
            var = tk.StringVar(value=self.cfg.get(key, dv))
            e = ttk.Entry(main, textvariable=var, show='*' if '密码' in key else None); e.grid(row=row, column=1, sticky='we')
            self.entries[key] = var
            self.config_widgets.append(e)
            var.trace_add('write', lambda *args, k=key, v=var: self.save_and_log_change(k, v.get()))
            row += 1

        # 调试模式
        self.debug = tk.BooleanVar(value=self.cfg.get('调试模式', False))
        debug_cb = ttk.Checkbutton(main, text='调试模式（打开浏览器可视化）', variable=self.debug)
        debug_cb.grid(row=row, column=0, columnspan=2, sticky='w', pady=(5,2))
        self.config_widgets.append(debug_cb)
        # 追踪调试模式改变
        self.debug.trace_add('write', lambda *args: self.save_and_log_change('调试模式', bool(self.debug.get())))
        row += 1

        # 添加验证码输入框
        ttk.Label(main, text='验证码').grid(row=row, column=0, sticky='e')
        self.verification_code_entry = ttk.Entry(main)
        self.verification_code_entry.grid(row=row, column=1, sticky='we')
        # 验证码不持久化到配置文件（短时输入），因此不加入 trace
        self.config_widgets.append(self.verification_code_entry)
        row += 1
        ttk.Label(main, text='非校园网用户需要填写验证码，点击”启动“即可获取验证码，退出填入即可', foreground='gray').grid(row=row, column=1, columnspan=2, sticky='we')
        row += 1
        # 启动按钮
        self.start_button = ttk.Button(main, text='启动', command=self.start)
        self.start_button.grid(row=row, column=0, columnspan=2, pady=(5,2))
        row += 1
        # 状态框
        status_f = ttk.LabelFrame(main, text='状态信息', padding=5)
        status_f.grid(row=row, column=0, columnspan=2, pady=(2,5), sticky='nsew')
        main.rowconfigure(row, weight=1)
        status_f.rowconfigure(0, weight=1)
        status_f.columnconfigure(0, weight=1)
        self.status = tk.Text(status_f, state='disabled')
        self.status.grid(row=0, column=0, sticky='nsew')
        sb = ttk.Scrollbar(status_f, command=self.status.yview)
        sb.grid(row=0, column=1, sticky='ns')
        self.status.config(yscrollcommand=sb.set)
        # 日志初始化
        setup_logging(self.status)
        # 版权信息
        row += 1
        ttk.Label(main, text='© 2025 NEU 监控助手', font=('微软雅黑', 12)).grid(row=row, column=0, columnspan=2, pady=(5,0))

    def start(self):
        for w in self.config_widgets:
            w.config(state='disabled')
        self.start_button.config(state='disabled')
        threading.Thread(target=self._run_monitor, daemon=True).start()

    def _run_monitor(self):
        # entries 中现在保存的是 StringVar，使用 get() 获取
        cfg = {k: v.get() for k, v in self.entries.items()}
        cfg.update({f'场地:{i}': v.get() for i, v in self.courts.items()})
        cfg.update({f'时段:{s}': v.get() for s, v in self.slots.items()})
        cfg['调试模式'] = self.debug.get()
        save_config(cfg)
        logging.info('开始监控')
        # 初始化浏览器并登录（在启动时需要验证码可能已填入）
        self.driver = init_driver(self.debug.get())
        verification_code = self.verification_code_entry.get()
        try:
            login_and_open_panel(self.driver, self.url, cfg['用户名'], cfg['登录密码'], verification_code)
        except Exception:
            logging.error('用户名和密码错误，或者登录失败，请检查后重试')
            for w in self.config_widgets:
                w.config(state='normal')
            self.start_button.config(state='normal')
            return

        # 准备监控参数，保存以便后续重启复用
        mail_cfg = [cfg['SMTP服务器'], int(cfg['端口']), cfg['邮箱'], cfg['SMTP密码'], cfg['收件']]
        courts = [i for i, v in self.courts.items() if v.get()]
        slots = [s for s, v in self.slots.items() if v.get()]
        base_interval = float(cfg['刷新间隔(s)'])
        max_retry = int(cfg['最大重试次数'])
        self.monitor_params = {
            'courts': courts,
            'slots': slots,
            'base_interval': base_interval,
            'max_retry': max_retry,
            'mail_cfg': mail_cfg
        }

        # 确保旧的 stop_event 被清除
        self._stop_event = threading.Event()
        # driver_getter 让监控线程在每次循环读取最新的 self.driver（这样 restart 会替换 self.driver）
        def driver_getter():
            return self.driver

        # 启动监控线程
        self.monitor_thread = threading.Thread(
            target=monitor_slots,
            args=(driver_getter, courts, slots, base_interval, max_retry, mail_cfg, self._stop_event),
            daemon=True
        )
        self.monitor_thread.start()

        # 安排 3 小时后触发重登录（使用 after 安排在主线程，实际重登录在单独线程执行）
        self.after(RELOGIN_INTERVAL, lambda: threading.Thread(target=self._perform_restart, daemon=True).start())

    def _perform_restart(self):
        logging.info('开始 3 小时到期自动重登录流程')
        # 首先通知监控线程停止
        self._stop_event.set()
        # 给监控线程一点时间退出（非阻塞等待）
        time.sleep(1.0)

        # 关闭旧浏览器
        try:
            if self.driver:
                self.driver.quit()
        except Exception as e:
            logging.warning(f'关闭旧浏览器时发生异常: {e}')
        self.driver = None

        # 重新初始化浏览器并登录（不使用验证码输入框）
        try:
            # 读取最新配置（可能用户在运行时修改了）
            cfg = load_config()
            logging.info('执行自动重登录（3小时）')
            self.driver = init_driver(self.debug.get())
            login_and_open_panel(self.driver, self.url, cfg.get('用户名',''), cfg.get('登录密码',''))
            logging.info('自动重登录成功')
        except Exception as e:
            logging.error(f'自动重登录失败: {e}')
            # 如果重登录失败，保留 stop_event 为已设置，稍后可以手动重启
            return

        # 重置 stop_event 并重启监控线程（复用之前保存的监控参数）
        self._stop_event = threading.Event()
        if self.monitor_params:
            params = self.monitor_params
            def driver_getter():
                return self.driver
            self.monitor_thread = threading.Thread(
                target=monitor_slots,
                args=(driver_getter, params['courts'], params['slots'], params['base_interval'], params['max_retry'], params['mail_cfg'], self._stop_event),
                daemon=True
            )
            self.monitor_thread.start()
            logging.info('重启监控线程完成')

        # 再次安排下一个 3 小时重登录
        self.after(RELOGIN_INTERVAL, lambda: threading.Thread(target=self._perform_restart, daemon=True).start())

    def restart(self):
        # 保留旧接口：立即异步触发一次重登录
        threading.Thread(target=self._perform_restart, daemon=True).start()

    def on_close(self):
        logging.info('程序关闭，退出监控')
        try:
            # 停止监控线程
            self._stop_event.set()
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        self.destroy()
        sys.exit(0)

if __name__ == '__main__':
    App().mainloop()