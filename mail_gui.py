import re
import os
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askdirectory,askopenfilename
from tkinter import ttk
import mail_send,mail_xlsx,threading
class gui():
    def __init__(self,root):
        self.root=root
        self.Gui()
    def Gui(self):
        self.botom = LabelFrame(self.root)
        self.botom.grid(row=1, column=0, padx=15, pady=2)
        self.mail_content_file_var = StringVar()
        self.mail_content_Label = Label(self.botom, text="邮件内容路径: ").grid(row=1, column=0, padx=15, pady=2, sticky="e")
        self.mail_content_Entry = Entry(self.botom, textvariable=self.mail_content_file_var,state='readonly', bd=2).grid(row=1, column=1, padx=15, pady=2)
        self.mail_content_Button = Button(self.botom, text="文件选择", command=self.select_content_file, relief=SOLID, bd=2).grid(row=1,
                                                                                                             column=2,
                                                                                                             padx=15,
                                                                                                             pady=2)
        self.mail_account_var = StringVar()
        self.mail_account_lable = Label(self.botom, text="账号批量导入").grid(row=2, column=0, padx=15, pady=2,
                                                                                       sticky="e")
        self.mail_account_Entry = Entry(self.botom, bd=2, textvariable=self.mail_account_var,state='readonly').grid(row=2, column=1, padx=15,
                                                                                             pady=2)
        self.mail_account_Button = Button(self.botom, text="文件选择", command=self.select_account_file, relief=SOLID, bd=2).grid(row=2,
                                                                                                             column=2,
                                                                                                             padx=15,
                                                                                                             pady=2)
        self.mail_username_var = StringVar()
        self.mail_username_lable = Label(self.botom, text="账号:").grid(row=3, column=0, padx=15, pady=2,
                                                                        sticky="e")
        self.mail_username_Entry = Entry(self.botom, bd=2, textvariable=self.mail_username_var).grid(row=3, column=1, padx=15,
                                                                        pady=2)
        self.mail_password_var = StringVar()
        self.mail_password_lable = Label(self.botom, text="密码:").grid(row=4, column=0, padx=15, pady=2,
                                                                      sticky="e")
        self.mail_username_Entry = Entry(self.botom, bd=2, textvariable=self.mail_password_var,show = '*').grid(row=4, column=1,
                                                                                                    padx=15,
                                                                                                    pady=2)
        self.btn = Button(self.botom, text="发送", relief=SOLID, bd=2, width=6, command=self.send)
        self.btn.grid(row=4, column=2, padx=15, pady=2, sticky="e")

    def send(self):
        def send_thread():
            messagebox.showinfo('邮件群发', '正在发送，请稍等！: ')
            send_mail = mail_send.email_send()
            read_xlsx = mail_xlsx.read_xlsx()
            if self.mail_content_file_var.get():
                content=read_xlsx(self.mail_content_file_var.get())
                if self.mail_account_var.get():
                    users=read_xlsx(self.mail_account_var.get())
                    try:
                        send_mail.send_more(content=content,users=users)
                    except Exception as e:
                        messagebox.showinfo('邮件群发', '发送失败！: '+str(e))
                    messagebox.showinfo('邮件群发', '发送完成！: ')
                else:
                    send_mail.username=self.mail_username_var.get()
                    send_mail.password=self.mail_password_var.get()
                    try:
                        send_mail.send_more(content=content)
                    except Exception as e:
                        messagebox.showinfo('邮件群发', '发送失败！: ' + str(e))
                    messagebox.showinfo('邮件群发', '发送完成！: ')
        thread = threading.Thread(target=send_thread)
        thread.start()

    def select_content_file(self):
        xlsx_ = askopenfilename(filetypes=[('Excel files', '*.xlsx;*.xls')])
        self.mail_content_file_var.set(xlsx_)

    def select_account_file(self):
        xlsx_ = askopenfilename(filetypes=[('Excel files', '*.xlsx;*.xls')])
        self.mail_account_var.set(xlsx_)
if __name__=="__main__":
    try:
        root=Tk()
        root.title("邮件群发")
        gui(root)
        root.mainloop()
    except Exception as e:
        print(e)
