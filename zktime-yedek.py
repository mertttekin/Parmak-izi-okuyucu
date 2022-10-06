from collections import UserDict, UserList
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

conn = None
# create ZK instance
zk = ZK('172.16.54.178', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
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
    worksheet.write("A1","PERSONEL" )
    worksheet.write("B1","TARİH" )
    worksheet.write("C1","GİRİŞ" )
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
            
            worksheet.write(akolon,name )
            worksheet.write(bkolon,tarih )
            worksheet.write(ckolon,zaman )
            satır+=1
        else:
            continue
        
    

    workbook.close()
    # excel_file = client.Dispatch("Excel.Application")
    # xl_sheets = excel_file.Workbooks.Open(r'C:\Users\Mert Tekin\Desktop\YazılımGeliştirme\zktime-aksiyon-script\Otomatik-gönderim.xlsx')
    # worksheets = xl_sheets.Worksheets[0]
    # worksheets.ExportAsFixedFormat(0, r'C:\Users\Mert Tekin\Desktop\YazılımGeliştirme\zktime-aksiyon-script\Otomatik-gönderim.pdf')
    
    smtp_server = "smtp.yandex.com.tr"
    port = 587  # For starttls
    sender_email = "ticketxx@aksiyonteknoloji.com"
    password = "T1xxxxxx"
    message = "test"
    receiver_email = "mert.tekin@aksiyonteknoloji.com"
    subject = "An email with attachment from Mert"
    body = "This is an email with attachment sent from Python"


    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

        # Add body to email
    message.attach(MIMEText(body, "plain"))

    filename = "test1.txt"  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as f:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(base64.b64encode(f.read()))

    # Encode file in ASCII characters to send by email    
    # encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.yandex.com.tr", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)





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
