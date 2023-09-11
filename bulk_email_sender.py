from PyQt5 import QtCore, QtWidgets, uic
from PyQt5.QtCore import *
import sys
import smtplib
import json
import os
import os.path
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr


def read_json_file():
    with open('./data/data.json','r') as f:
        data = json.load(f)
    return data

def write_json_file(data):
    with open('./data/data.json','w') as f:
        json.dump(data, f, indent=4)
    return data

def update_json(new_data):
    with open('./data/data.json','r') as f:
        data = json.load(f)
    data.update(new_data)
    with open('./data/data.json','w') as f:
        json.dump(data, f, indent=4)
    return data

json_data = read_json_file()

class MainUi(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainUi, self).__init__()
        uic.loadUi('./data/main_window.ui', self)
        self.setWindowTitle("Mail Sender")
        self.lineEdit_smtp_server.setText(str(json_data['smtp_server']))
        self.lineEdit_smtp_port.setText(str(json_data['smtp_port']))
        self.lineEdit_account_name.setText(str(json_data['account_name']))
        self.lineEdit_account_pass.setText(str(json_data['account_pass']))
        self.lineEdit_from_email.setText(str(json_data['sender_mail']))
        self.lineEdit_from_name.setText(str(json_data['sender_name']))
        self.lineEdit_signature.setText(str(json_data['signature']))
        self.lineEdit_delay_time.setText(str(json_data['delay_time']))
        self.textEdit_cc_mails.setPlainText(str(json_data['cc_mails']))
        self.attachments_path.setVisible(False)
        self.label_last_sent_info.setVisible(False)
        self.label_progress.setVisible(False)

        self.start_btn.clicked.connect(self.start_thread)
        self.stop_btn.clicked.connect(self.stop_thread)
        self.save_btn.clicked.connect(self.save_info)
        self.attachments_btn.clicked.connect(self.open_attachment_folder)
        self.show()

    def save_info(self):
        new_data = {
            "smtp_server": str(self.lineEdit_smtp_server.text()),
            "smtp_port": str(self.lineEdit_smtp_port.text()),
            "account_name": str(self.lineEdit_account_name.text()),
            "account_pass": str(self.lineEdit_account_pass.text()),
            "sender_mail": str(self.lineEdit_from_email.text()),
            "sender_name": str(self.lineEdit_from_name.text()),
            "cc_mails": str(self.textEdit_cc_mails.toPlainText()),
            "signature": str(self.lineEdit_signature.text()),
            "delay_time": float(self.lineEdit_delay_time.text())
        }
        update_json(new_data)
        os.execl(sys.executable, sys.executable, *sys.argv)

    def open_attachment_folder(self):
        self.upload_file = QtWidgets.QFileDialog.getExistingDirectory(self, "Klasör Seç", "C:/")
        if self.upload_file:
            file_list = "\n".join(os.listdir(self.upload_file))
            self.attachments_path.setText(file_list)
            self.attachments_path.setVisible(True)

    def start_thread(self):
        self.thread = StartThreadClass(parent=None,index=0)
        self.thread.smtp_server = self.lineEdit_smtp_server.text()
        self.thread.smtp_port = self.lineEdit_smtp_port.text()
        self.thread.account_name = self.lineEdit_account_name.text()
        self.thread.account_pass = self.lineEdit_account_pass.text()
        self.thread.sender_mail = self.lineEdit_from_email.text()
        self.thread.sender_name = self.lineEdit_from_name.text()
        self.thread.signature_path = self.lineEdit_signature.text()
        self.thread.delay_time = self.lineEdit_delay_time.text()
        self.thread.subject = self.lineEdit_mail_subject.text()
        self.thread.to_mails = self.textEdit_to_mails.toPlainText()
        self.thread.cc_mails = self.textEdit_cc_mails.toPlainText()
        self.thread.message_body = self.textEdit_message_body.toPlainText()
        if self.attachments_path.text() != "Folder Path":
            self.thread.upload_path = self.upload_file
        self.thread.file_list = self.attachments_path.text()
        self.attachments_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.label_last_sent_info.setVisible(True)
        self.label_progress.setText('Processing...')
        self.label_progress.setVisible(True)
        self.thread.start()
        self.thread.last_sent_signal.connect(self.label_last_sent_info.setText)
        self.thread.timing_signal.connect(self.lcdNumber.display)
        self.thread.process_signal.connect(self.label_progress.setText)

    def stop_thread(self):
        self.attachments_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.attachments_path.setText('Folder Path')
        self.attachments_path.setVisible(False)
        self.thread.stop()
        

class StartThreadClass(QtCore.QThread):
    last_sent_signal = QtCore.pyqtSignal(str)
    timing_signal = QtCore.pyqtSignal(int)
    process_signal = QtCore.pyqtSignal(str)
    def __init__(self, parent: None,index=0):
        super(StartThreadClass, self).__init__(parent)
        self.index = index
        self.is_running = True
        self.control_check = False

    def run(self):
        try:
            to_mails = self.to_mails.split(',')
            cc_mails = self.cc_mails.split(',')
            for receiver in to_mails:
                message = MIMEMultipart()
                message['Subject'] = self.subject
                message['From'] = formataddr((self.sender_name, self.sender_mail))
                message['To'] = formataddr(('', receiver))
                message['CC'] = formataddr(('',','.join(cc_mails)))
                to_s = [receiver] + cc_mails
                message_body = self.message_body.replace('\n','<br>')
                text_content = MIMEText(message_body, 'html', 'utf-8')
                message.attach(text_content)
                if self.file_list != "Folder Path":
                    attachments = self.file_list.split('\n')
                    for fpath in attachments:
                        with open(self.upload_path+'/'+fpath, 'rb') as f:
                            part = MIMEApplication(f.read())
                            part.add_header('Content-Disposition', 'attachment',
                                            filename=os.path.basename(fpath))
                            message.attach(part)

                with open(self.signature_path, 'r') as f:
                    signature = f.read()
                signature_content = MIMEText(signature.encode('utf-8'), 'html', 'utf-8')
                message.attach(signature_content)

                try:
                    server = smtplib.SMTP_SSL(self.smtp_server,int(self.smtp_port))
                    server.ehlo()
                    server.login(self.account_name, self.account_pass)
                    server.sendmail(self.sender_mail, to_s, message.as_string())
                    self.last_sent_signal.emit(str(receiver))
                except Exception as e:
                    print(e)
                finally:
                    server.quit()
                for i in range(int(float(self.delay_time)*60),-1,-1):
                    self.timing_signal.emit(i)
                    time.sleep(1)
                
            self.process_signal.emit('Finished')
        except Exception as e:
            self.process_signal.emit(e)
            

    def stop(self):
        self.is_running = False
        self.terminate()



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainUi()
    sys.exit(app.exec_())