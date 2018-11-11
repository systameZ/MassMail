import smtplib,time,random,threading
import email.mime.multipart
import email.mime.text
"""
    基本发送

"""
class email_send(object):
    def __init__(self,username=None,password=None,to_eamil=None,subject=None,content=None):
        self.msg=email.mime.multipart.MIMEMultipart()
        self.to = '' if  to_eamil is None else to_eamil
        #self.from_mail = ''if  username is None else username
        self.subject = '' if  subject is None else subject
        self.content = '' if  content is None else content
        self.username='' if  username is None else username
        self.password=''if  password is None else password
        self.lock=threading.Lock()

    def send(self):
        smtp = smtplib.SMTP()
        try:
            smtp.connect(self.splite_name(),'25')
            smtp.login(user=self.username,password=self.password)
        except Exception as e:
            print('登录失败！: ',e)
            smtp.quit()
            self.msg = email.mime.multipart.MIMEMultipart()
        try:

            smtp.sendmail(self.msg['from'], self.msg['to'], str(self.msg))
        except Exception as se:
            print('发送失败！: ',se)
            self.msg = email.mime.multipart.MIMEMultipart()
        smtp.quit()
        self.msg = email.mime.multipart.MIMEMultipart()
    def send_qq(self):
        smtp = smtplib.SMTP()
        try:
            server=smtplib.SMTP_SSL(self.splite_name(),'465')
            server.login(user=self.username,password=self.password)
        except Exception as e:
            print('登录失败！: ',e)
            server.quit()
            self.msg = email.mime.multipart.MIMEMultipart()
        try:

            server.sendmail(self.msg['from'], self.msg['to'], str(self.msg))
        except Exception as se:
            print('发送失败！: ',se)
            self.msg = email.mime.multipart.MIMEMultipart()
        server.quit()
        self.msg = email.mime.multipart.MIMEMultipart()
    def splite_name(self):#分割出服务器地址
        if self.username:
            mail_server=self.username.split('@')[1]
        return 'smtp.'+mail_server
    def  send_more(self,content=None,users=None):
        if users:
            user_list = []
            for user in users:
                user_list.append(user)
        if content:
            touser_list=[]
            for tomail in content:
                touser_list.append(tomail)
        def thread_send():
            #while(touser_list):
                if content :
                        if  self.lock.acquire():
                            if users:
                                self.username=self.make_random(user_list)
                                self.password=users[self.username]
                            to_user=touser_list.pop()
                            self.msg['to']=to_user
                            self.content=content[to_user][0]
                            self.msg['subject']=content[to_user][1]
                            self.msg['from'] = self.username
                            self.text = email.mime.text.MIMEText(self.content)
                            self.msg.attach(self.text)
                            if self.splite_name()=='smtp.qq.com':
                                self.send_qq()
                            else:
                                self.send()
                            self.lock.release()
                            time.sleep(3)


        threads = []
        while threads or touser_list:
            for thread in threads:
                if not thread.is_alive():
                    threads.remove(thread)
            while len(threads) < 3 and touser_list:
                thread = threading.Thread(target=thread_send)
                thread.setDaemon(True)
                thread.start()
                threads.append(thread)
            time.sleep(1)
    def make_random(self,users):#随机取出一个账号，返回账号密码
        ran = random.sample(users,1)
        return ran[0]


if __name__=='__main__':
    pass
    sendemail=email_send('test@qq.com','xigafaqscozycbbe','test@163.com','学习email','收到')
    sendemail.send()#xigafaqscozycbbe
