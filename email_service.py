# -*- coding: utf-8 -*-
import os
import time
import json
import logging
import threading
import smtplib
import ssl
import socket
import zipfile
from pathlib import Path
from email.message import EmailMessage

# =========================================================
# 📧 邮件服务配置
# =========================================================

# 1. 默认管理员邮箱 (备用)
#    当配置文件存在但内容为空/无效时，将发送到此地址
DEFAULT_ADMIN_EMAIL = "winlogger@189.cn" 

# 2. 配置文件名称 (纯文本格式)
EMAIL_CONFIG_FILENAME = "wll.config.ini"

# 3. 本地发送记录文件名
HISTORY_FILENAME = "wll.archive.json"

# 4. 发件人池 (主备轮询机制)
SENDER_POOL = [
    # [Primary] 主发件人: 中国电信 189 邮箱
    # Host: smtp.189.cn | Port: 465 (SSL)
    ("smtp.189.cn", 465, "winlogger@189.cn", "Bf*9My@5e@3Oh(3J"),

    # [Backup] 备用发件人 (可留空)
    # ("smtp.gmail.com", 465, "backup@gmail.com", "BACKUP_PASSWORD"),
]

# 5. 策略配置
INITIAL_DELAY_SECONDS = 300  # 启动后 5 分钟检查
RETRY_INTERVAL_SECONDS = 600 # 失败后 10 分钟重试
MAX_RETRIES_PER_SESSION = 3  # 最大重试次数

# =========================================================

class EmailSender:
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.config_file = self.base_path / EMAIL_CONFIG_FILENAME
        self.history_file = self.base_path / HISTORY_FILENAME
        
        # 1. 初始化时确保配置文件存在 (若不存在则创建 'do not send')
        self._ensure_config_exists()
        
        self.sent_files = self._load_history()
        self.lock = threading.Lock()
        
        self.dirs_to_scan = ["Hardware", "Events"]

    def _ensure_config_exists(self):
        """确保配置文件存在，默认禁用邮件发送"""
        if not self.config_file.exists():
            try:
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    f.write("do not send\n")
                    f.write("# [Instructions]\n")
                    f.write("# Default: 'do not send' (Email disabled)\n")
                    f.write("# To enable: Replace the first line with target email (e.g. boss@189.cn)\n")
                logging.info(f"Created default config file: {self.config_file}")
            except Exception as e:
                logging.error(f"Failed to create default config: {e}")

    def _load_history(self):
        if self.history_file.exists():
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return set(json.load(f))
            except: return set()
        return set()

    def _save_history(self):
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(list(self.sent_files), f)
        except Exception as e:
            logging.error(f"Failed to save email history: {e}")

    def get_receiver(self):
        """
        读取配置文件:
        - "do not send" -> 返回 None (不发)
        - 空文件/无效内容 -> 返回 DEFAULT_ADMIN_EMAIL (发给默认)
        - 有效邮箱 -> 返回该邮箱
        """
        target = DEFAULT_ADMIN_EMAIL
        
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    if not lines: return target
                    
                    first_line = lines[0].strip()
                    if first_line.lower() == "do not send":
                        logging.info("Email disabled by config.")
                        return None
                    
                    if "@" in first_line:
                        target = first_line
            except Exception as e:
                logging.warning(f"Error reading config: {e}")
        
        return target

    def check_internet(self):
        try:
            socket.create_connection(("114.114.114.114", 53), timeout=3)
            return True
        except: return False

    def get_device_name(self):
        try: return socket.gethostname()
        except: return "UnknownDevice"

    def scan_files_to_send(self):
        files_to_send = []
        for log_dir in self.dirs_to_scan:
            dir_path = self.base_path / log_dir
            if not dir_path.exists(): continue
            for f in dir_path.glob("*.xlsx"):
                if f.name not in self.sent_files:
                    files_to_send.append(f)
        return files_to_send

    def create_zip_archive(self, files, device_name, date_range_str):
        # 使用英文半角括号
        zip_name = f"Logs_{device_name}({date_range_str}).zip"
        zip_path = self.base_path / "cache" / zip_name
        try:
            zip_path.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for file_path in files:
                    zf.write(file_path, arcname=file_path.name)
            return zip_path
        except Exception as e:
            logging.error(f"Failed to create zip: {e}")
            return None

    def send_batch(self):
        receiver = self.get_receiver()
        if not receiver: return True # 配置为不发送，视为任务完成

        if not self.check_internet():
            logging.warning("No internet. Email skipped.")
            return False

        files = self.scan_files_to_send()
        if not files:
            logging.info("No new logs.")
            return True

        # 日期范围
        dates = set()
        for f in files:
            parts = f.name.split('_')
            if len(parts) > 1: dates.add(parts[1])
        sorted_dates = sorted(list(dates))
        date_range = f"{sorted_dates[0]}"
        if len(sorted_dates) > 1: date_range += f"~{sorted_dates[-1]}"

        device_name = self.get_device_name()

        # 打包
        logging.info(f"Compressing {len(files)} files...")
        zip_path = self.create_zip_archive(files, device_name, date_range)
        if not zip_path: return False

        # 构建邮件
        subject = f"Logs_{device_name}({date_range})"
        body = f"{device_name}'s Logs"

        success = False
        last_error = ""

        for host, port, user, password in SENDER_POOL:
            if "REPLACE" in password or "your_" in user: continue
            try:
                msg = EmailMessage()
                msg['Subject'] = subject
                msg['From'] = user
                msg['To'] = receiver
                msg.set_content(body)

                with open(zip_path, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='zip', filename=zip_path.name)

                logging.info(f"Sending via {host}...")
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(host, port, context=context, timeout=60) as smtp:
                    smtp.login(user, password)
                    smtp.send_message(msg)
                
                logging.info(f"Sent to {receiver}")
                success = True
                break
            except Exception as e:
                last_error = str(e)
                logging.error(f"Failed via {host}: {e}")
                continue

        try:
            if zip_path.exists(): os.remove(zip_path)
        except: pass

        if success:
            with self.lock:
                for f in files: self.sent_files.add(f.name)
                self._save_history()
            return True
        else:
            logging.error(f"All senders failed: {last_error}")
            return False

def _email_worker(base_path, stop_event):
    sender = EmailSender(base_path)
    logging.info(f"Email scheduler waiting {INITIAL_DELAY_SECONDS}s...")
    if stop_event.wait(INITIAL_DELAY_SECONDS): return

    retry_count = 0
    while not stop_event.is_set():
        try:
            if sender.send_batch():
                logging.info("Email task completed.")
                break
            else:
                retry_count += 1
                if retry_count >= MAX_RETRIES_PER_SESSION:
                    logging.warning("Max retries reached.")
                    break
                logging.info(f"Retrying in {RETRY_INTERVAL_SECONDS}s...")
                if stop_event.wait(RETRY_INTERVAL_SECONDS): break
        except Exception as e:
            logging.error(f"Email worker error: {e}")
            break

def start_email_service(base_path, stop_event):
    t = threading.Thread(target=_email_worker, args=(base_path, stop_event), daemon=True)
    t.start()
    return t
