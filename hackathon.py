import binascii
import threading
import nfc
import time
import csv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formatdate
from os.path import basename
import smtplib

##############################################

json_path=''#jsonのパス
sheet_path=''#csv保存パス

#メール送信内容
from_add = ''#送信元
password=''#送信元のメアドのパスワード
to_add =''#送信先
subject = '入退出確認' #件名
body = ''#本文内容

#送信時間
H=20
M=30

##############################################


service_code = 0x09CB
scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(
        json_path, 
        scope)
gc = gspread.authorize(credentials)
wb = gc.open('oecu_iclog').sheet1

def getNextRow():
    values_list = wb.col_values(1)
    emptyrow = str(len(values_list)+1)
    return emptyrow
    
def on_connect_nfc1(tag):
    with open(sheet_path,'a') as f:
      if isinstance(tag, nfc.tag.tt3.Type3Tag):
        try:
            #print('  ' + '\n  '.join(tag.dump()))
            row = getNextRow()
            sc = nfc.tag.tt3.ServiceCode(service_code >> 6 ,service_code & 0x3f)
            bc1 = nfc.tag.tt3.BlockCode(0,service=0)
            bc2 = nfc.tag.tt3.BlockCode(1,service=0)
            data = tag.read_without_encryption([sc],[bc1,bc2])
            
            now_date = datetime.now()
            dir = now_date.strftime('%Y')+"/" + now_date.strftime('%m')+"/"+ now_date.strftime("%d")+"/" + now_date.strftime('%H')+":" + now_date.strftime('%M')+":" + now_date.strftime('%S')
            
            sid =  str(data[2:10])
            oecuname=sid[12:20]
            name_str = ''.join(oecuname)
            name_list = [name_str]
            name_str2 = ''.join(dir)
            name_list2 = [name_str2]
            enterstr=''.join('入室')
            enter=[enterstr]
            name_list3 = [name_list] + [name_list2] +[enter]
            
            print(name_list3)
            writer = csv.writer(f)
            writer.writerow(name_list3)
            
            wb.update_acell('A'+row,oecuname)
            wb.update_acell('B'+row,dir)
            wb.update_acell('C'+row , '入室')

        except Exception as e:
            print ("error1: %s" % e)    
            
def on_connect_nfc2(tag):
    with open(sheet_path,'a') as f:
      if isinstance(tag, nfc.tag.tt3.Type3Tag):
        try:
            #print('  ' + '\n  '.join(tag.dump()))
            row = getNextRow()
            sc = nfc.tag.tt3.ServiceCode(service_code >> 6 ,service_code & 0x3f)
            bc1 = nfc.tag.tt3.BlockCode(0,service=0)
            bc2 = nfc.tag.tt3.BlockCode(1,service=0)
            data = tag.read_without_encryption([sc],[bc1,bc2])
            
            now_date = datetime.now()
            dir = now_date.strftime('%Y')+"/" + now_date.strftime('%m')+"/"+ now_date.strftime("%d")+"/" + now_date.strftime('%H')+":" + now_date.strftime('%M')+":" + now_date.strftime('%S')
            
            sid =  str(data[2:10])
            oecuname=sid[12:20]
            name_str = ''.join(oecuname)
            name_list = [name_str]
            name_str2 = ''.join(dir)
            name_list2 = [name_str2]
            enterstr2=''.join('退室')
            enter2=[enterstr2]
            name_list3 = [name_list] + [name_list2] + [enter2]
            
            print(name_list3)
            writer = csv.writer(f)
            writer.writerow(name_list3)
            f.close()
            
            wb.update_acell('A'+row,oecuname)
            wb.update_acell('B'+row,dir)
            wb.update_acell('C'+row ,'退室')


        except Exception as e:
            print ("error1: %s" % e)
    
def create_message(from_addr, to_addr, subject, body):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Date'] = formatdate()
    msg.attach(MIMEText(body))
    
    path = sheet_path
    with open(path, "rb") as f:
        part = MIMEApplication(f.read(),Name=basename(path))
    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(path)
    msg.attach(part)
    return msg

def send_mail(from_addr, to_addr, body_msg):
    smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpobj.ehlo()
    smtpobj.starttls()
    smtpobj.ehlo()
    smtpobj.login(from_addr,password)
    smtpobj.sendmail(from_addr, to_addr, body_msg.as_string())
    smtpobj.close()

def main1():
    clf = nfc.ContactlessFrontend('usb:001:004')
    while True:
            clf.connect(rdwr={'on-connect': on_connect_nfc1})
            time.sleep(3)
        
def main2():
    clf = nfc.ContactlessFrontend('usb:001:003')
    while True:
        clf.connect(rdwr={'on-connect': on_connect_nfc2})
        time.sleep(3)
        
def timeinfo():
    while True:
        now_date = datetime.now()
        now_time = now_date.strftime('%H')+ now_date.strftime('%M')+ now_date.strftime('%S')
        send_timming = now_date.replace(hour=H).strftime('%H')+now_date.replace(minute=M).strftime('%M')+ now_date.replace(second=0).strftime('%S')
        if int(now_time)-int(send_timming)==0:
            msg = create_message(from_add, to_add, subject, body)
            #print(msg)
            send_mail(from_add,to_add, msg)
            print("メール送信")
            #print(now_time)
            break
        
if __name__ == "__main__":
    try:
        t1=threading.Thread(target=main1)
        t2=threading.Thread(target=main2)
        t3=threading.Thread(target=timeinfo)
        t1.start()
        t2.start()
        t3.start()
    except KeyboardInterrupt:
        print("\nFinish")
