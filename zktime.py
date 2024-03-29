from collections import UserDict, UserList
from email.mime.application import MIMEApplication
from posixpath import basename
from zk import ZK, const
import xlsxwriter
from datetime import datetime
import smtplib, ssl
import email, smtplib, ssl
from win32com import client
import base64
import os

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# def config_read():
#     global alicimail
#     filename="external.config"
#     contents = open(filename).read()
#     config = eval(contents)
#     alicimail = config['alici']
#     print(alicimail)
    

conn = None
# create ZK instance
zk = ZK('172.16.54.xxx', port=4300, timeout=5, password=0, force_udp=False, ommit_ping=False)
try:
    # connect to device
    conn = zk.connect()
    # disable device, this method ensures no activity on the device while the process is run
    conn.disable_device()
    # another commands will be here!
    # Example: Get All Users
    users = conn.get_users()
    for user in users:
        privilege = 'User'
        if user.privilege == const.USER_ADMIN:
            privilege = 'Admin'
        print ('+ UID #{}'.format(user.uid))
        print ('  Name       : {}'.format(user.name))
        print ('  Privilege  : {}'.format(privilege))
        print ('  Password   : {}'.format(user.password))
        print ('  Group ID   : {}'.format(user.group_id))
        print ('  User  ID   : {}'.format(user.user_id))

    userss = conn.read_sizes()
    print(conn)
    print(userss)
    print ('  record  : {}'.format(conn.users))
    print ('  record  : {}'.format(conn.fingers))
    print ('  record  : {}'.format(conn.records))
    print ('  record  : {}'.format(conn.users_cap))
    print ('  record  : {}'.format(conn.fingers_cap))

    
    attendances = conn.get_attendance()
    print(attendances)
    print(type(attendances))
    global usersozluk
    usersozluk={ 
        1:"Name" ,
        2:"Erkan Y",
        3:"Fatih O",
        4:"Dilek A",
        5:"Ozcan M",
        6:"Aksiyon",
        7:"Volkan S",
        8:"Eylem P",
        9:"Oyku D Y",
        10:"Semih K",
        11:"None",
        12:"Ercan Y",
        13:"Oktay S",
        14:"Ergün E",
        15:"Yiğit O",
        16:"Hakan P",
        17:"Serife A",
        18:"None",
        19:"Ersin E",
        20:"Mert T"                        
                            }
    print(usersozluk)
    print(len(attendances))
    print(attendances[0])
    value1 = str(attendances[0])
    print(type(value1))
    print(value1[14:15])
    

    global tarih
    tarih = datetime.today().strftime('%Y-%m-%d')
    print (tarih)
    # excelismi = "%s-tarihi-kayıtlar.xlsx" %(tarih)
    excelismi = "Otomatik-gönderim.xlsx"
    global workbook
    workbook = xlsxwriter.Workbook(excelismi)
    worksheet = workbook.add_worksheet()
    cell_format = workbook.add_format({'bold': True, 'font_color': 'black'})
    cell_format.set_bg_color('blue')
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    worksheet.set_row(0, 25)
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 20)
    worksheet.write("A1","PERSONEL", cell_format )
    worksheet.write("B1","TARİH", cell_format )
    worksheet.write("C1","GİRİŞ", cell_format )
    # set the width of the column

    # worksheet.row_dimensions[1].height = 20
    # worksheet.column_dimensions['B'].width = 20
    global satır
    global stun
    satır = 2
    stun = 1
    # name = usersozluk[2]
    # print(name)

    for i in range(len(attendances)):
        value1 = str(attendances[i])
        userid = int(value1[14:16].strip())
        name = usersozluk[userid]
        zaman = value1[29:38].strip()
        basmatarihi = value1[18:29].strip()
        if tarih == basmatarihi:    
            akolon = "A%s"%(satır) 
            bkolon = "B%s"%(satır) 
            ckolon = "C%s"%(satır) 
            if zaman >= "09:00:00":
                cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
                worksheet.write(akolon,name , cell_format )
                worksheet.write(bkolon,tarih, cell_format )
                worksheet.write(ckolon,zaman, cell_format )
            else:
                cell_format = workbook.add_format({'bold': True, 'font_color': 'green'})
                worksheet.write(akolon,name, cell_format )
                worksheet.write(bkolon,tarih, cell_format )
                worksheet.write(ckolon,zaman, cell_format )
            satır+=1
        else:
            continue
        
    

    workbook.close()
    # excel_file = client.Dispatch("Excel.Application")
    # xl_sheets = excel_file.Workbooks.Open(r'C:\Users\Mert Tekin\Desktop\YazılımGeliştirme\zktime-aksiyon-script\Otomatik-gönderim.xlsx')
    # worksheets = xl_sheets.Worksheets[0]
    # worksheets.ExportAsFixedFormat(0, r'C:\Users\Mert Tekin\Desktop\YazılımGeliştirme\zktime-aksiyon-script\Otomatik-gönderim.pdf')
    # config_read()
    smtp_server = "smtp.yandex.com.tr"
    port = 587  # For starttls
    sender_email = "ticket@aksiyonteknoloji.com"
    password = "xxxxx"
    message = "test"
    # receiver_email = "erkan.yetis@aksiyonteknoloji.com"
    receiver_email = "xx@aksiyonteknoloji.com,xxxsxx@aksiyonteknoloji.com,sxxxxs@aksiyonteknoloji.com,xxxx@aksiyonteknoloji.com,satinalma@aksiyonteknoloji.com,xxxxm@aksiyonteknoloji.com"
    receiver_email_mert = "mert.tekin@aksiyonteknoloji.com"
    toaddr = ['erkxxx@aksiyonteknoloji.com','mert.tekin@aksiyonteknoloji.com']
    cc = ['xxx@aksiyonteknoloji.com','xxxxx@aksiyonteknoloji.com','xxxx@aksiyonteknoloji.com']
    subject = "An email with attachment from Mert"
    body = "{} tarihli parmak okumalar.".format(tarih)


    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ', '.join(toaddr)
    message["Subject"] = subject
    message["Cc"] = ', '.join(cc)
    body = MIMEText(body,'plain')
    message.attach(body)

    filename = "Otomatik-gönderim.xlsx"  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as f:
        attecment = MIMEApplication(f.read(),Name=basename(filename))
        attecment['Content'] = 'attachment;filename="{}"'.format(basename(filename)) 

    message.attach(attecment)
    # Encode file in ASCII characters to send by email    
    # encoders.encode_base64(part)

    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.yandex.com.tr", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, (toaddr+cc), text)





    # Create a secure SSL context
    context = ssl.create_default_context()

    # Try to log in to server and send email
    # try:
    #     server = smtplib.SMTP(smtp_server,port)
    #     server.ehlo() # Can be omitted
    #     server.starttls(context=context) # Secure the connection
    #     server.ehlo() # Can be omitted
    #     server.login(sender_email, password)
    #     server.sendmail(sender_email, receiver_email, message)
    # except Exception as e:
    #     # Print any error messages to stdout
    #     print(e)
    # finally:
    #     server.quit() 
    
        
    # for attendance in conn.live_capture():
    #     if attendance is None:
    #         # implement here timeout logic
    #         pass
    #     else:
    #         print (attendance) # Attendance object
    # # Clear attendances records
    # conn.clear_attendance()


    # Test Voice: Say Thank You
    conn.test_voice()
    # re-enable device after all commands already executed
    conn.enable_device()
except Exception as e:
    print ("Process terminate : {}".format(e))
finally:
    if conn:
        conn.disconnect()
