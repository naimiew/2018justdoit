# -*- coding: utf-8 -*-
import os,re
import cx_Oracle
import time
from datetime import datetime
import requests,json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import parseaddr,formataddr
from bs4 import BeautifulSoup

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
username = "username"
passwd = "passwd"
host = "host"
port = "port"
sid = "sid"

dsn = cx_Oracle.makedsn(host, port, sid)
con = cx_Oracle.connect(username, passwd, dsn)
cursor = con.cursor()
print 'orcale connect success!!!!!!!!!!!!!'

sqlhead = "SELECT MERCHANT_ORDER_NO FROM x where to_char(CREATE_DATE_TIME,'yyyy-mm-dd HH24:MI:SS')>'"
#T日支付成功，状态ordersts = z
sqltail = "' and PRODUCT_ID = 'y' and ordersts = 'z'"
#生成当天sql时间条件2018-09-07 00:00:00
sqlstrptime = datetime.now().strptime(datetime.now().strftime('%b-%d-%Y'),'%b-%d-%Y')
#拼接sql语句，datetime转str
sqlstr = sqlhead+str(sqlstrptime)+sqltail
#print sqlstr
sqlgo = sqlstr
cursor.execute(sqlgo)

result = cursor.fetchall()
#print("Total: " + str(cursor.rowcount))

order_list = []
MERCHANT_ORDER_NO = ''
orig_order_no_body = ''
order_no = ''
for row in result: #清洗数据库中数据
	mixstr = str(row)
	MERCHANT_ORDER_NO = re.findall("\d+",mixstr)[0]
	order_list.append(str(MERCHANT_ORDER_NO))
print order_list

cursor.close()
con.close()

for orig_order_no in order_list:
	installment_cancel_url = "installment_cancel_url"
	installment_cancel_time = int(time.mktime(datetime.now().timetuple()))
	order_no = 'order_no'+str(installment_cancel_time)
	data = {'merchant_id':'merchant_id',
	                   'orig_order_no':orig_order_no,
	                   'order_no':order_no,
	                   'note':'',
	                   'env':'env',
	                   'submit':'submit'
	                   }
	print data
	r = requests.post(installment_cancel_url, data)
	time.sleep(1)
	print r.content
	soup = BeautifulSoup(r.content, "lxml")
	tag = soup.body
	orig_order_no_body += '<br>MERCHANT_ORDER_NO:' + orig_order_no +'<br>'+ 'cancel:' + order_no + '<br>' + tag.string +'<br>~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

# 格式化邮件地址
def formatAddr(s):
	name, addr = parseaddr(s)
	return formataddr((Header(name, 'utf-8').encode(), addr))


def sendMail(body):
	smtp_server = 'smtp.163.com'
	from_mail = 'from_mail'
	mail_pass = 'mail_pass'
	to_mail = ['huhy@reapal.com']
	# 构造一个MIMEMultipart对象代表邮件本身
	msg = MIMEMultipart()
	# Header对中文进行转码
	msg['From'] = formatAddr('测试环境撤销定时 <%s>' % from_mail).encode()
	msg['To'] = ','.join(to_mail)
	msg['Subject'] = Header('Total' + str(cursor.rowcount), 'utf-8').encode()
	msg.attach(MIMEText(body, 'html', 'utf-8'))
	try:
		s = smtplib.SMTP()
		s.connect(smtp_server, "25")
		s.login(from_mail, mail_pass)
		s.sendmail(from_mail, to_mail, msg.as_string())  # as_string()把MIMEText对象变成str
		s.quit()
	except smtplib.SMTPException as e:
		print "Error: %s" % e


if __name__ == "__main__":
	body = orig_order_no_body
	sendMail(body)
	

