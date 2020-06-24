# 群发录用通知
# xxxxx 授权码
import email
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.utils import formataddr
import smtplib
import xlrd
import docx
import os
from wordGen import genWord
# xxxx 授权码

sink = r'xxxx'  # PDF存储目录

# 读入参会者信息
data = xlrd.open_workbook(r'xxxx.xlsx')
table = data.sheet_by_name('info')

name = table.col_values(0)[1:]   # 姓名
paper = table.col_values(4)[1:]  # 论文题目
mailBox = table.col_values(3)[1:]  # 电子邮箱
direction = table.col_values(6)[1:]  # 分论坛名
form = table.col_values(8)[1:]   # 参会形式

# name = name[44:]
# mailBox = mailBox[44:]
sender = 'xxxx.com'
userName = 'xxxx'
pwd = 'xxxx'
host = 'smtp.126.com'
port = 25

# sender = 'xxxx.com'
# userName = 'kay'
# pwd = 'xxxxx'
# host = 'smtp.163.com'
# port = 25


# PDF文件名
def fileName(name):
    return '录用通知' + '-' + name + '.pdf'


def sendMails(name, email):
    # 正文
    body = "<p style=\"margin:0;\">尊敬的"+name+":</p>\
<p style=\"margin:0;\">&nbsp; &nbsp; &nbsp; &nbsp;谨代表xxxx，欢迎您参加xxxxx！</p>\
<p style=\"margin:0;\">&nbsp; &nbsp; &nbsp; &nbsp;附件是录用通知，请查收。</p>\
<p style=\"margin:0;\">&nbsp; &nbsp; &nbsp; &nbsp;参会前还请您配合完成:</p>\
<p style=\"margin:0;\">&nbsp; &nbsp; &nbsp; &nbsp;1 请尚未上传照片的代表，于<b>6月21日</b>前将<b>个人照片(证件照/生活照)</b>发送至本邮箱，以方便会议准备；</p>\
<p style=\"margin:0;\">&nbsp; &nbsp; &nbsp; &nbsp;2 请及时扫码加入<b>论坛微信群</b>，以获取最新信息与会议安排。</p>\
<p style=\"margin:0;\">&nbsp; &nbsp; &nbsp; &nbsp;感谢您对本论坛的支持与配合！</p>\
<p style=\"margin:0;\">祝好</p>\
<p style=\"margin:0;\">xxxxxx</p>\
<p> <img src=\"cid:image1\"></p>"
    recip = email

    msg = MIMEMultipart()
    msg.set_charset('utf-8')
    msg['From'] = formataddr(pair=(userName, sender))    # 这个用来设置发件人名字的  ps:真的没力气再研究了...

    # msg.add_header('From', sender)
    msg.add_header('To', recip)
    msg.add_header('Subject', '录用通知')

    # 文字内容,HTML格式
    msg.attach(MIMEText(body, 'html', _charset='utf-8'))
    # 添加图片
    # 这样是加在附件里
    # with open('weChat.png', 'rb') as fp:
    #     msg.attach(MIMEImage(fp.read()))
    fp = open('weChat3.png', 'rb')
    msgImage = MIMEImage(fp.read())
    msgImage.add_header('Content-ID', '<image1>')
    msg.attach(msgImage)

    # 附件
    attach = MIMEBase('pdf', 'pdf')
    pdfPath = os.path.join(sink, fileName(name))
    with open(pdfPath, 'rb') as pf:
        # 加上必要的头信息:
        attach.add_header('Content-Disposition', 'attachment', filename=fileName(name))
        attach.add_header('Content-ID', '<0>')
        attach.add_header('X-Attachment-Id', '0')

        attach.set_payload(pf.read())
        encoders.encode_base64(attach)
        msg.attach(attach)

    server.send_message(msg)



# 发邮件
i = 0
server = smtplib.SMTP(host, port)
server.starttls()
server.login(sender, pwd)
# sendMails('xx', 'xxxx.com')

for pName, pEmail in zip(name, mailBox):
# for j in range(50):
    i += 1
    # sendMails('xx', 'xxxx.com')
    sendMails(pName, pEmail)
    print(pName + '发送完成\n')
    # 主要是因为发现发11封服务器就需要重连
    if i % 10 == 0:
        server.quit()
        server = smtplib.SMTP(host, port)
        server.starttls()
        server.login(sender, pwd)
server.quit()
