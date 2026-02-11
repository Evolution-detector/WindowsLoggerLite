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

# 1. 默认管理员邮箱 (源码预设，当没有配置文件时发给它)
DEFAULT_ADMIN_EMAIL = "winlogger@189.cn" 

# 2. 配置文件名称 (纯文本格式，同目录下)
EMAIL_CONFIG_FILENAME = "wll.config.ini"

# 3. 本地发送记录文件名
HISTORY_FILENAME = "wll.archive.json"

# 4. 发件人池 (主备轮询机制)
SENDER_POOL = [
    # [Primary] 主发件人: 中国电信 189 邮箱
    # Host: smtp.189.cn | Port: 465 (SSL)
    ("smtp.189.cn", 465, "winlogger@189.cn", "Bf*9My@5e@3Oh(3J"),

    # [Backup] 备用发件人 (预留位置，暂时为空)
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
        self.history_file = self.base_path / HISTORY_FILENAME
        self.sent_files = self._load_history()
        self.lock = threading.Lock()
        
        # 对应主程序中的日志文件夹名称
        self.dirs_to_scan = ["Hardware", "Events"]

    def _load_history(self):
        """加载已发送文件列表"""
        if self.history_file.exists():
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return set(json.load(f))
            except: return set()
        return set()

    def _save_history(self):
        """保存发送记录"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(list(self.sent_files), f)
        except Exception as e:
            logging.error(f"Failed to save email history: {e}")

    def get_receiver(self):
        """
        读取配置文件 (wll.config.ini)
        逻辑: 纯文本模式。读取第一行有效内容作为收件人。
        """
        config_path = self.base_path / EMAIL_CONFIG_FILENAME
        target = DEFAULT_ADMIN_EMAIL
        
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    # 读取全部内容，去除首尾空格
                    content = f.read().strip()
                    
                    # 判断特殊指令
                    if content.lower() == "do not send":
                        logging.info("Email feature disabled by config ('do not send').")
                        return None
                    
                    # 简单的有效性检查
                    if "@" in content:
                        target = content
            except Exception as e:
                logging.warning(f"Error reading email config: {e}")
        
        return target

    def check_internet(self):
        """连通性测试 (尝试连接 114 DNS)"""
        try:
            socket.create_connection(("114.114.114.114", 53), timeout=3)
            return True
        except: return False

    def get_device_name(self):
        """获取计算机名 (Hostname)"""
        try:
            return socket.gethostname()
        except:
            return "UnknownDevice"

    def scan_files_to_send(self):
        """扫描尚未发送的 Excel 日志"""
        files_to_send = []
        for log_dir in self.dirs_to_scan:
            dir_path = self.base_path / log_dir
            if not dir_path.exists(): continue
            
            for f in dir_path.glob("*.xlsx"):
                if f.name not in self.sent_files:
                    files_to_send.append(f)
        return files_to_send

    def create_zip_archive(self, files, device_name, date_range_str):
        """
        打包成 ZIP
        文件名格式: Logs_{Device Name}({日期范围}).zip
        """
        zip_name = f"Logs_{device_name}({date_range_str}).zip"
        zip_path = self.base_path / "cache" / zip_name
        
        try:
            # 确保 cache 目录存在
            zip_path.parent.mkdir(parents=True, exist_ok=True)
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for file_path in files:
                    # arcname 是压缩包内的文件名 (不带绝对路径)
                    zf.write(file_path, arcname=file_path.name)
            return zip_path
        except Exception as e:
            logging.error(f"Failed to create zip file: {e}")
            return None

    def send_batch(self):
        receiver = self.get_receiver()
        if not receiver: return True # 配置为不发送，视为任务完成

        if not self.check_internet():
            logging.warning("No internet connection. Email skipped.")
            return False

        files = self.scan_files_to_send()
        if not files:
            logging.info("No new logs to email.")
            return True

        # --- 生成日期范围字符串 ---
        dates = set()
        for f in files:
            # 文件名格式: UUID_2025-01-01_...
            parts = f.name.split('_')
            if len(parts) > 1: dates.add(parts[1])
        
        sorted_dates = sorted(list(dates))
        date_range = f"{sorted_dates[0]}"
        if len(sorted_dates) > 1:
            date_range += f"~{sorted_dates[-1]}"

        device_name = self.get_device_name()

        # --- 📦 打包逻辑 ---
        logging.info(f"Compressing {len(files)} files into zip archive...")
        zip_path = self.create_zip_archive(files, device_name, date_range)
        if not zip_path:
            return False # 打包失败，稍后重试

        # --- 构建邮件 ---
        # 主题: Logs_{Device Name}({日期范围})
        subject = f"Logs_{device_name}({date_range})"
        # 正文: {Device Name}'s Logs
        body = f"{device_name}'s Logs"

        success = False
        last_error = ""

        # --- 轮询发件人池 ---
        for host, port, user, password in SENDER_POOL:
            if "REPLACE" in password or "your_" in user: continue

            try:
                msg = EmailMessage()
                msg['Subject'] = subject
                msg['From'] = user
                msg['To'] = receiver
                msg.set_content(body)

                # 添加 ZIP 附件
                with open(zip_path, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='zip', filename=zip_path.name)

                logging.info(f"Attempting to send email via {host}...")
                context = ssl.create_default_context()
                # 189邮箱和其他标准SSL邮箱都使用 SMTP_SSL
                with smtplib.SMTP_SSL(host, port, context=context, timeout=60) as smtp:
                    smtp.login(user, password)
                    smtp.send_message(msg)
                
                logging.info(f"Email sent successfully to {receiver}")
                success = True
                break
            except Exception as e:
                last_error = str(e)
                logging.error(f"Failed to send via {host}: {e}")
                continue

        # 清理临时 ZIP 文件
        try:
            if zip_path.exists():
                os.remove(zip_path)
        except: pass

        if success:
            with self.lock:
                for f in files: self.sent_files.add(f.name)
                self._save_history()
            return True
        else:
            logging.error(f"All senders failed. Last error: {last_error}")
            return False

def _email_worker(base_path, stop_event):
    sender = EmailSender(base_path)
    
    logging.info(f"Email scheduler waiting {INITIAL_DELAY_SECONDS}s...")
    # 启动延迟
    if stop_event.wait(INITIAL_DELAY_SECONDS): return

    retry_count = 0
    while not stop_event.is_set():
        try:
            if sender.send_batch():
                logging.info("Email task completed.")
                break # 发送成功，本次运行使命结束
            else:
                retry_count += 1
                if retry_count >= MAX_RETRIES_PER_SESSION:
                    logging.warning("Max email retries reached.")
                    break
                logging.info(f"Retrying email in {RETRY_INTERVAL_SECONDS}s...")
                # 失败等待
                if stop_event.wait(RETRY_INTERVAL_SECONDS): break
        except Exception as e:
            logging.error(f"Email worker error: {e}")
            break

def start_email_service(base_path, stop_event):
    """主程序调用的入口函数"""
    t = threading.Thread(target=_email_worker, args=(base_path, stop_event), daemon=True)
    t.start()
    return t
