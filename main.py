import json
import re
import smtplib
import ssl
import time
import sys
import threading
import os
import requests
from datetime import datetime, date, timedelta
from email.mime.text import MIMEText

from playwright.sync_api import sync_playwright
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

# ================= 全局配置 (完全对齐第一版脚本) =================
AUTH_URL = "https://raw.githubusercontent.com/DDDDDDKing/lib/main/auth.json"
STATE_FILE = "user_config.json"

# GMX 邮箱发件配置
SMTP_HOST = "smtp.gmx.com"
SMTP_PORT = 465
SMTP_SENDER = "carsonqtes_emley@gmx.com"
SMTP_AUTH_CODE = "O5vohhrk99"

PUSHPLUS_TOKEN = "5f1adcce3e70435a9946059cb3ccdb98"

# 业务参数
START_TIME_VAL = "08:00"
END_TIME_VAL = "22:00"
FLOOR_AREA = {2: 76, 3: 77, 4: 78, 5: 79}
BASE_DATE = date(2026, 1, 6)
BASE_SEGMENT_4F = 1589715
FLOOR_OFFSET = {2: -9064, 3: -4532, 4: 0, 5: 4532}

BOOK_RECORD_URL = "https://libzw.csu.edu.cn/user/index/book"

# 浏览器驱动路径
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(BASE_DIR, "pw_browsers")

class LibraryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("校方资源自动化预约终端 (双重安全版)")
        self.root.geometry("620x850")
        self.is_running = False
        self.stop_event = threading.Event()
        self.setup_ui()
        self.load_config()

    def setup_ui(self):
        style = ttk.Style()
        style.configure("Header.TLabel", font=("微软雅黑", 14, "bold"), foreground="#FFFFFF", background="#2C3E50")
        style.configure("Warn.TLabel", font=("微软雅黑", 9), foreground="#E74C3C", background="#2C3E50")
        
        header = tk.Frame(self.root, bg="#2C3E50", height=80)
        header.pack(fill="x")
        ttk.Label(header, text="选座助手", style="Header.TLabel").pack(pady=5)
        ttk.Label(header, text="提示：00:00-06:00 系统维护期程序将自动休眠", style="Warn.TLabel").pack()

        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill="both", expand=True)

        # 1. 账号配置
        login_lf = ttk.LabelFrame(main_frame, text=" 认证配置 ")
        login_lf.pack(fill="x", pady=5)
        opts = {'padx': 5, 'pady': 5}
        ttk.Label(login_lf, text="学号:").grid(row=0, column=0, sticky="e", **opts)
        self.ent_user = ttk.Entry(login_lf, width=20)
        self.ent_user.grid(row=0, column=1, **opts)
        ttk.Label(login_lf, text="密码:").grid(row=0, column=2, sticky="e", **opts)
        self.ent_pass = ttk.Entry(login_lf, show="●", width=20)
        self.ent_pass.grid(row=0, column=3, **opts)
        ttk.Label(login_lf, text="通知邮箱:").grid(row=1, column=0, sticky="e", **opts)
        self.ent_email = ttk.Entry(login_lf, width=45)
        self.ent_email.grid(row=1, column=1, columnspan=3, **opts)

        # 2. 任务设置
        task_lf = ttk.LabelFrame(main_frame, text=" 任务参数 ")
        task_lf.pack(fill="x", pady=5)
        ttk.Label(task_lf, text="预约日期:").grid(row=0, column=0, sticky="e", **opts)
        self.cb_date = ttk.Combobox(task_lf, values=["今天", "明天"], state="readonly", width=10)
        self.cb_date.set("明天")
        self.cb_date.grid(row=0, column=1, **opts)
        ttk.Label(task_lf, text="目标层级:").grid(row=0, column=2, sticky="e", **opts)
        self.cb_floor = ttk.Combobox(task_lf, values=["2", "3", "4", "5"], state="readonly", width=10)
        self.cb_floor.set("4")
        self.cb_floor.grid(row=0, column=3, **opts)
        ttk.Label(task_lf, text="编号(如1,2-10):").grid(row=1, column=0, sticky="e", **opts)
        self.ent_seats = ttk.Entry(task_lf, width=45)
        self.ent_seats.insert(0, "4000-4050")
        self.ent_seats.grid(row=1, column=1, columnspan=3, **opts)

        # 3. 启动定时
        timer_lf = ttk.LabelFrame(main_frame, text=" 定时启动控制 ")
        timer_lf.pack(fill="x", pady=5)
        self.var_immediate = tk.BooleanVar(value=True)
        ttk.Checkbutton(timer_lf, text="立即启动", variable=self.var_immediate).grid(row=0, column=0, **opts)
        
        self.cb_h = ttk.Combobox(timer_lf, values=[f"{i:02d}" for i in range(24)], width=4, state="readonly")
        self.cb_h.set("21")
        self.cb_h.grid(row=0, column=1, sticky="e")
        ttk.Label(timer_lf, text="时").grid(row=0, column=2, sticky="w")
        self.cb_m = ttk.Combobox(timer_lf, values=[f"{i:02d}" for i in range(60)], width=4, state="readonly")
        self.cb_m.set("59")
        self.cb_m.grid(row=0, column=3, sticky="e")
        ttk.Label(timer_lf, text="分").grid(row=0, column=4, sticky="w")
        self.cb_s = ttk.Combobox(timer_lf, values=[f"{i:02d}" for i in range(60)], width=4, state="readonly")
        self.cb_s.set("59")
        self.cb_s.grid(row=0, column=5, sticky="e")
        ttk.Label(timer_lf, text="秒").grid(row=0, column=6, sticky="w")

        # 4. 操作台
        self.btn_start = ttk.Button(main_frame, text="🚀 开始执行任务", command=self.start_task)
        self.btn_start.pack(fill="x", pady=5)
        self.btn_stop = ttk.Button(main_frame, text="⛔ 强制终止", state="disabled", command=self.stop_task)
        self.btn_stop.pack(fill="x")
        
        self.log_area = scrolledtext.ScrolledText(main_frame, height=15, font=("Consolas", 9), state="disabled")
        self.log_area.pack(fill="both", expand=True, pady=10)
        self.log_area.tag_config("SUCCESS", foreground="green")
        self.log_area.tag_config("ERROR", foreground="red")
        self.log_area.tag_config("WAIT", foreground="#D35400")

    def log(self, msg, level="NORMAL"):
        self.log_area.config(state="normal")
        t = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert("end", f"[{t}] {msg}\n", level)
        self.log_area.see("end")
        self.log_area.config(state="disabled")

    def load_config(self):
        if os.path.exists(STATE_FILE):
            try:
                with open(STATE_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.ent_user.insert(0, data.get("username", ""))
                    self.ent_pass.insert(0, data.get("password", ""))
                    self.ent_email.insert(0, data.get("email", ""))
            except: pass

    def save_config(self):
        data = {"username": self.ent_user.get(), "password": self.ent_pass.get(), "email": self.ent_email.get()}
        with open(STATE_FILE, "w", encoding="utf-8") as f: json.dump(data, f)

    def start_task(self):
        self.save_config()
        self.is_running = True
        self.stop_event.clear()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        threading.Thread(target=self.run_logic, daemon=True).start()

    def stop_task(self):
        self.stop_event.set()
        self.log("任务已手动停止", "ERROR")

    def run_logic(self):
        user = self.ent_user.get().strip()
        pwd = self.ent_pass.get().strip()
        
        # 权限确认日志
        self.log("正在与云端服务器确认用户授权权限...", "WAIT")
        try:
            r = requests.get(AUTH_URL, timeout=5)
            auth_data = r.json()
            if not auth_data.get("enabled", False) or user not in auth_data.get("allowed_users", []):
                self.log("❌ 权限确认失败：您的凭证未在校方允许名单内", "ERROR")
                self.root.after(0, lambda: self.btn_start.config(state="normal"))
                return
            self.log("✅ 权限确认通过，允许访问校方资源系统", "SUCCESS")
        except:
            self.log("⚠ 权限确认超时，请检查网络连接", "ERROR")
            self.root.after(0, lambda: self.btn_start.config(state="normal"))
            return

        # 定时等待
        if not self.var_immediate.get():
            target_t = f"{self.cb_h.get()}:{self.cb_m.get()}:{self.cb_s.get()}"
            self.log(f"⏳ 任务挂起，等待启动时间: {target_t}", "WAIT")
            while datetime.now().strftime("%H:%M:%S") < target_t and not self.stop_event.is_set():
                time.sleep(0.1)

        while not self.stop_event.is_set():
            # 维护期休眠
            if 0 <= datetime.now().hour < 6:
                self.log("🌙 处于系统维护期，程序自动休眠，06:00将唤醒", "WAIT")
                while datetime.now().hour < 6 and not self.stop_event.is_set():
                    time.sleep(30)
                continue

            try:
                with sync_playwright() as p:
                    browser = p.chromium.launch(headless=False)
                    page = browser.new_page()
                    
                    days = 0 if self.cb_date.get() == "今天" else 1
                    target_date = date.today() + timedelta(days=days)
                    target_url = self.build_url(target_date, int(self.cb_floor.get()))
                    
                    page.goto(target_url)
                    self.try_login(page, user, pwd)
                    
                    try: page.click("#nav-date > a.btn.btn-default.active.area_day", timeout=3000)
                    except: pass

                    target_seats = self.parse_seats(self.ent_seats.get())
                    
                    while not self.stop_event.is_set():
                        if datetime.now().hour == 0: break
                        
                        page.reload()
                        page.wait_for_timeout(500)
                        
                        seats = page.query_selector_all("li.seat.ava-icon")
                        for s in seats:
                            s_data = json.loads(s.get_attribute("data-data") or "{}")
                            s_no = str(s_data.get("no", ""))
                            
                            if s_no in target_seats:
                                self.log(f"🎯 正在尝试请求资源编号: {s_no}")
                                s.click()
                                page.wait_for_timeout(200)
                                page.keyboard.press("Enter")
                                page.wait_for_timeout(200)
                                page.keyboard.press("Enter")
                                
                                # 双重校验：后端记录确认 + 座位号一致性校验
                                if self.verify_and_compare(page, user, s_no):
                                    self.stop_event.set()
                                    break
                                else:
                                    self.log(f"❌ 校验失败：后端未发现编号 {s_no} 的成功记录，返回重试")
                                    page.goto(target_url)
                        
                        if self.stop_event.is_set(): break
                        time.sleep(0.8)
                    browser.close()
            except Exception as e:
                self.log(f"⚠ 系统通讯异常: {e}", "ERROR")
                time.sleep(3)

        self.root.after(0, lambda: self.btn_start.config(state="normal"))

    def try_login(self, page, u, p):
        btn = page.wait_for_selector("a.login-btn.login_click", timeout=10000)
        btn.click()
        page.fill("#username", u)
        page.fill('input[name="passwordText"]', p)
        page.keyboard.press("Enter")
        page.wait_for_selector("li.seat", timeout=15000)

    def verify_and_compare(self, page, user, target_s_no):
        """核心校验：跳转后端并比对座位号"""
        try:
            self.log("🔍 正在核实后端任务记录...", "WAIT")
            page.goto(BOOK_RECORD_URL, timeout=10000)
            page.wait_for_selector("#menu_table", timeout=10000)
            
            rows = page.query_selector_all("#menu_table tbody tr")
            for row in rows:
                content = row.inner_text()
                # 只有状态为成功或使用中，且内容包含我们点击的座位号才算真成功
                if ("预约成功" in content or "使用中" in content) and (target_s_no in content):
                    tds = row.query_selector_all("td")
                    info = {
                        "id": tds[0].inner_text().strip(),
                        "space": tds[1].inner_text().strip().replace("\n", ""),
                        "start": tds[2].inner_text().strip(),
                        "status": tds[4].inner_text().strip()
                    }
                    self.send_final_email(user, info)
                    return True
            return False
        except: return False

    def send_final_email(self, user, info):
        """完全按照 monitor_final.py 的 ssl 发信逻辑"""
        msg_text = f"🎉 资源预约最终确认成功！\n\n凭证：{user}\n资源名称：{info['space']}\n生效时间：{info['start']}\n流水单号：{info['id']}\n最终状态：{info['status']}"
        self.log(msg_text, "SUCCESS")
        
        email_to = self.ent_email.get().strip()
        if not email_to: return

        try:
            mime = MIMEText(msg_text, "plain", "utf-8")
            mime["Subject"] = "【系统通知】预约任务成功核实"
            mime["From"] = SMTP_SENDER
            mime["To"] = email_to

            # 严格使用 ssl.create_default_context()
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
                server.login(SMTP_SENDER, SMTP_AUTH_CODE)
                server.send_message(mime)
            self.log("📧 确认通知邮件已成功投递", "SUCCESS")
        except Exception as e:
            self.log(f"📧 邮件模块故障: {e}", "ERROR")

        try: requests.post("https://www.pushplus.plus/send", json={"token": PUSHPLUS_TOKEN, "title": "任务圆满完成", "content": msg_text}, timeout=5)
        except: pass

    def build_url(self, target_date, floor):
        day_diff = (target_date - BASE_DATE).days
        segment = BASE_SEGMENT_4F + day_diff + FLOOR_OFFSET[floor]
        area = FLOOR_AREA[floor]
        return f"https://libzw.csu.edu.cn/web/seat3?area={area}&segment={segment}&day={target_date}&startTime={START_TIME_VAL}&endTime={END_TIME_VAL}"

    def parse_seats(self, text):
        res = set()
        for p in text.replace("，", ",").split(","):
            if "-" in p:
                try:
                    s, e = p.split("-")
                    for i in range(int(s), int(e)+1): res.add(str(i))
                except: pass
            elif p.strip().isdigit(): res.add(p.strip())
        return res

if __name__ == "__main__":
    root = tk.Tk()
    app = LibraryApp(root)
    root.mainloop()
