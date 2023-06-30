
"""
=-=-=-=-=--=-=-=- ℹ INFORMAÇÕES -=-=-=-=-=-=-=-=-=-=-=
AUTOR: Christian Lima Freitag 📚
PROJETO: Labld-tools - Email automatizado 🎨
-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=--=-=-=-=-=

"""

from cgitb import text
from time import sleep
import win32com.client as win32
import smtplib
import email.message as em
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


import json
from tkinter import END, OptionMenu, PhotoImage, StringVar, Tk, BOTH, Frame, Button, Text, Label, messagebox, Menu, PhotoImage, HORIZONTAL, END, filedialog
from tkinter.ttk import Style, Progressbar
import os
import sys
import xlwings as xw
import threading


class Exmail(Frame):

    def __init__(self, controller):
        super().__init__()
        self.index_frame = 0
        self.ct = controller
        self.initUI(controller)  # construtor chamando função para instanciar 🧮
        config = self.onLoad()  # pegart configurações do json
        # Inicializando variaveis de configuração

    def send_email(self, emailto, msge, assunto, op_server):

        if(op_server == "Outlook (Windows 8,10)"):

            try:
                self.outlook_email.To = emailto
                self.outlook_email.Subject = assunto
                self.outlook_email.HTMLBody = msge
                self.outlook_email.Send()
                print('Email para '+emailto+'enviado 📧')
            except:
                # self.error = True
                messagebox.showerror(
                    title="Erro de envio", message="Ocorreu um erro ao enviar emails com o servidor outlook.")
        else:

            message = MIMEMultipart()

            message['From'] = self.email_server
            message['To'] = emailto.replace(" ", "")
            message['Subject'] = assunto

            message.attach(MIMEText(msge, 'html'))

            session = smtplib.SMTP(self.endereço_server)
            session .starttls()

            session.login(message['From'].replace(" ", "").replace(
                "\n", ""), self.senha_server.replace(" ", "").replace("\n", ""))

            text = message.as_string()

            try:
                session.sendmail(message['From'], [message['To']],
                                 text)

                print('Email para '+emailto+' enviado 📧 de '+self.email_server)

                session.quit()
            except:
                messagebox.showerror(
                    title="Erro de envio", message="Ocorreu um erro ao enviar o email para "+str(emailto))
                session.quit()
                # self.error = True

    def onLoad(self):
        config = None

        # puxando dados do arquivo de configuração json
        if(os.path.exists("config.json")):
            f = open("config.json", "r", encoding="utf-8")
            content = f.read()
            config = json.loads(content)
        else:
            f = open("config.json", "w", encoding="utf-8")
            f.write("{'op_server': 'Outro', 'endereco_server': 'gmail.com: 587', 'email_server': '', 'senha_server': '', 'excel_coluna_email': '', 'excel_pagina': '', 'excel_colunas_de preenchimento': '','format_mesage':''}".replace("'", '"'))
            f.close()
            config = json.loads(
                "{'op_server': 'Outro', 'endereco_server': 'gmail.com: 587', 'email_server': '', 'senha_server': '', 'excel_coluna_email': '', 'excel_pagina': '', 'excel_colunas_de preenchimento': '','format_mesage':''}".replace("'", '"'))

        # carregando os dados do json de configuração nas variaveis
        self.op_server = config["op_server"]
        self.endereço_server = config["endereco_server"]
        self.email_server = config["email_server"]
        self.senha_server = config["senha_server"]
        self.coluna_email = config["excel_coluna_email"]
        self.pagina_name = config["excel_pagina"]
        self.colunas_de_preenchimento = config["excel_colunas_de preenchimento"]
        self.format_mensagem = config["format_mesage"]

        # inserção ao abrir para os campos de serviço de email
        self.service_email.insert("1.0", self.email_server)
        self.service_password.insert("1.0", self.senha_server)
        self.service_name.insert(
            "1.0", str(" ".join(self.endereço_server.split(":")[0].split(".")[1:])).replace(" ", "."))
        self.service_port.insert("1.0", str(
            self.endereço_server.split(":")[1]).replace(" ", ""))
        self.mensagem_file_name_field.insert("1.0", self.format_mensagem)
        self.coluna_email_excel.insert("1.0", self.coluna_email)
        self.coluna_preenchimento.insert("1.0", self.colunas_de_preenchimento)
        self.pagina_excel.insert("1.0", self.pagina_name)

        # modificações para menu se serviçoes de email
        self.variable.set(self.services[self.services.index(self.op_server)])
        self.onChangeServer(self.op_server)

    def switch_frame(self, n):
        if(self.op_server == ""):
            self.variable.set(self.services[0])
            self.onChangeServer("Gmail")
        self.frames_list[self.index_frame].pack_forget()
        self.frames_list[n].pack(fill="both", expand=1)
        self.index_frame = n

    # Abre diretório atual para pegar o arquivo do excel e retorna o endereço do arquivo
    def extract_from_excel(self, f):

        try:
            app = xw.App(visible=False)
            wb = xw.Book(f)
            sheet = wb.sheets[self.pagina_name]
            num_row = sheet.range(str(self.coluna_email+'2')).end('down').row
            self.label_descri.config(
                text="Foram encontrados "+str(num_row)+" endereços.")
            self.label_descri.place(x=20, y=120)
            self.colunas_de_preen_list = self.colunas_de_preenchimento.split(
                ";")
            self.c_d_p_aux = []

            self.emails_list = sheet.range(
                str(self.coluna_email+'2'), str(self.coluna_email+str(num_row))).value
            for i in self.colunas_de_preen_list:
                self.c_d_p_aux.append(sheet.range(
                    str(i+'2'), str(i+str(num_row))).value)
            wb.close()
            app.kill()
        except:
            wb.close()
            app.kill()

    def onOpen(self):
        # Função pra abrir o arquivo 📁
        ftypes = [('Arquivos excel', '*.xlsx'), ('Arquivos Html',
                                                 '*.htm'), ('Arquivo de texto', '*.txt')]
        dlg = filedialog.askopenfilename(
            initialdir=os.getcwd(), title="Select file", filetypes=ftypes)
        fl = dlg
        if fl != '':
            if(self.index_frame == 0):
                self.excel_file_name_field.delete("1.0", END)
                self.excel_file_name_field.insert('1.0', str(fl))
            elif(self.index_frame == 3):
                self.mensagem_file_name_field.delete('1.0', END)
                self.mensagem_file_name_field.insert('1.0', str(fl))

        if(self.index_frame == 0):
            self.extract_from_excel(fl)

    def onChangeServer(self, event):
        self.op_server = event
        if(event == "Outro"):

            self.label_email_service.place(x=20, y=55)
            self.service_email.place(x=80, y=55, width=150)
            self.label_password_service.place(x=250, y=55)
            self.service_password.place(x=300, y=55, width=120)

            self.label_descri_service.place(x=20, y=85)
            self.service_name.place(x=160, y=85, width=85)
            self.service_port_label.place(x=250, y=85)
            self.service_port.place(x=390, y=85, width=80)
        elif(event == "Gmail"):
            self.label_email_service.place(x=20, y=55)
            self.service_email.place(x=80, y=55, width=150)
            self.label_password_service.place(x=250, y=55)
            self.service_password.place(x=300, y=55, width=120)
        else:
            self.label_descri_service.place_forget()
            self.service_name.place_forget()
            self.service_port_label.place_forget()
            self.service_port.place_forget()
            self.label_email_service.place_forget()
            self.service_email.place_forget()
            self.label_password_service.place_forget()
            self.service_password.place_forget()

    def save_configs(self):
        f = open("config.json", "w", encoding="utf-8")
        jso = str("{'op_server': '"+self.op_server+"', 'endereco_server': '"+self.endereço_server+"', 'email_server': '"+self.email_server+"', 'senha_server': '"+self.senha_server+"', 'excel_coluna_email': '" +
                  self.coluna_email+"', 'excel_pagina': '"+self.pagina_name+"', 'excel_colunas_de preenchimento': '"+self.colunas_de_preenchimento+"','format_mesage':'"+self.format_mensagem.replace('"', "'")+"'}").replace("'", '"')
        f.write(jso)
        f.close()

    def save_service(self):
        if(self.op_server == "Outro"):
            self.email_server = self.service_email.get("1.0", "end-1c")
            self.senha_server = self.service_password.get("1.0", "end-1c")
            self.endereço_server = "smtp."+str(self.service_name.get("1.0", "end-1c").replace(
                " ", ""))+": "+str(self.service_port.get("1.0", "end-1c"))

        elif(self.op_server == "Gmail"):
            self.email_server = self.service_email.get("1.0", "end-1c")
            self.senha_server = self.service_password.get("1.0", "end-1c")
            self.endereço_server = 'smtp.gmail.com: 587'

        if(self.index_frame == 2):
            self.coluna_email = self.coluna_email_excel.get("1.0", "end-1c")
            self.pagina_name = self.pagina_excel.get("1.0", "end-1c")
            self.colunas_de_preenchimento = self.coluna_preenchimento.get(
                "1.0", "end-1c")
        self.format_mensagem = self.mensagem_file_name_field.get(
            "1.0", "end-1c")

        self.save_configs()
        messagebox.showinfo("Pronto", "Configuração salva!")
        self.switch_frame(0)

    def start_thread(self,):
        x = threading.Thread(target=self.start)
        x.start()

    def check_requirements(self,):

        if(self.op_server == "Gmail" or self.op_server == "Outro"):
            if(self.service_email == "" or self.format_mensagem == "" or self.assunto_field.get("1.0", "end-1c") == "" or self.excel_file_name_field.get("1.0", "end-1c") == "" or self.service_password == "" or self.service_name == "" or self.service_port == "" or self.endereço_server == "" or self.coluna_email == "" or self.pagina_name == ""):
                # self.error = False
                return False
            else:
                return True

        else:
            if(self.format_mensagem == "" or self.assunto_field.get("1.0", "end-1c") == "" or self.excel_file_name_field.get("1.0", "end-1c") == ""):
                return False
            else:
                try:
                    outlook = win32.Dispatch('outlook.application')
                    self.outlook_email = outlook.CreateItem(0)
                    # self.error = False
                    return True
                except:
                    messagebox.showerror(
                        title="Erro de envio", message="Erro para se conectar com servidor outlook.")
                    # self.error = True
                    return False

    def start(self,):
        op = messagebox.askyesno(
            title="Confirmação", message="Ao continuar não sera possível parar o processo. Deseja continuar?")

        if(op):
            if(self.check_requirements()):

                self.progressbar.place(x=240, y=115)
                self.progressbar['value'] = 0
                self.label_descri.config(
                    text="Começando..")

                file_msg = open(self.format_mensagem,
                                "r", encoding="utf-8")
                msg = file_msg.read()

                for email in range(len(self.emails_list)):
                    msg_aux = msg
                    for i in range(len(self.colunas_de_preen_list)):
                        if(self.c_d_p_aux[i][email] != None and self.c_d_p_aux[i][email] != ""):
                            msg_aux = msg_aux.replace(
                                "<#!=" + self.colunas_de_preen_list[i]+"=!#>",  self.c_d_p_aux[i][email].replace("\n", "<br>"))
                        else:
                            msg_aux = msg_aux.replace(
                                "<#!=" + self.colunas_de_preen_list[i]+"=!#>",  "")
                    try:
                        self.label_descri.config(
                            text=f'{email} de {len(self.emails_list)} e-mails enviados.')
                        sleep(int(self.intervalo_field.get("1.0", "end-1c")))
                        self.ct.update_idletasks()

                        self.progressbar['value'] += (100 /
                                                      len(self.emails_list))
                        x = threading.Thread(target=self.send_email, args=(self.emails_list[email], msg_aux, self.assunto_field.get(
                            "1.0", "end-1c"), self.op_server))
                        x.start()

                        # self.send_email(self.emails_list[email], msg_aux, self.assunto_field.get(
                        #    "1.0", "end-1c"), self.op_server)

                    except:
                        messagebox.showerror(
                            title="Erro de envio", message="Ocorreu um erro durante o envio dos emails, contate o suporte.")
                        # self.error = True
                        break

                self.progressbar['value'] = 0
                self.progressbar.place_forget()
                self.label_descri.config(text="")
            else:
                messagebox.showerror(
                    title="Erro de envio", message="Preencha todas as configurações/campos referentes ao envio do email.")

    def initUI(self, controller):  # Função de criação de elementos 🎨

        # ?- Configurações ⚙
        s = Style()
        s.theme_use("clam")

        # ?  - Inicializando frame inical 🐾
        # frame inicial onde os emelemntos vão ficar em cima
        frame_main = Frame(self,)

        frame_main.pack(expand=True, fill=BOTH)  # instanciando frame
        frame_servers = Frame(self,)
        frame_modelo_mensagem = Frame(self)
        frame_extrac_excel = Frame(self,)
        frame_about = Frame(self,)
        self.frames_list = [frame_main, frame_servers,
                            frame_extrac_excel, frame_modelo_mensagem, frame_about]
        self.pack(fill="both", expand=1)

        # ? Menu bar
        menubar = Menu(self)
        file = Menu(menubar, tearoff=1)
        file.add_command(label="Serviço de email",
                         command=lambda: self.switch_frame(1))
        file.add_command(label="Modelo de mensagem",
                         command=lambda: self.switch_frame(3))
        file.add_command(label="Extração do excel",
                         command=lambda: self.switch_frame(2))
        menubar.add_cascade(label="Configurações", menu=file)
        edit = Menu(menubar, tearoff=0)
        edit.add_command(label="Sobre", command=lambda: self.switch_frame(4))
        menubar.add_cascade(label="Exmail", menu=edit)
        controller.config(menu=menubar)

        # arquivo excel

        label_descri = Label(frame_main, text=(
            "Selecione o arquivo do excel: "), foreground="black", font=('Helvetica 9'))
        label_descri.place(x=20, y=10)

        escolher_buttom = Button(
            frame_main, text="Escolher", command=self.onOpen, height=1)
        escolher_buttom.place(x=400, y=32, width=80)

        self.excel_file_name_field = Text(
            frame_main, height="1", foreground="black", borderwidth=1, bg="white")
        self.excel_file_name_field.place(x=20, y=35, width=380)

        # intervalo

        label_descri = Label(frame_main, text=(
            "Intervalo (segundos): "), foreground="black", font=('Helvetica 9'))
        label_descri.place(x=20, y=70)

        self.intervalo_field = Text(
            frame_main, height="1", foreground="black", borderwidth=1,)
        self.intervalo_field.insert("1.0", 2)
        self.intervalo_field.place(x=160, y=70, width=40)

        label_assunto = Label(frame_main, text=(
            "Assunto: "), foreground="black", font=('Helvetica 9'))
        label_assunto.place(x=220, y=70)

        self.assunto_field = Text(
            frame_main, height="1", foreground="black", borderwidth=1)
        self.assunto_field.place(x=280, y=70, width=200)

        startbt = Button(
            frame_main, text="Começar", height=1, command=self.start_thread)
        startbt.place(x=360, y=110, width="120")

        self.progressbar = Progressbar(
            frame_main, orient=HORIZONTAL, length=100, mode='determinate')

        # info

        self.label_descri = Label(frame_main, text=(
            "Foram encontrados 235 email."), foreground="grey", font=('Helvetica 9'))

        # voltar buttons

        back_buttom1 = Button(frame_servers, text="Voltar",
                              height=1, command=lambda: self.switch_frame(0))
        back_buttom1.place(x=20, y=110)

        back_buttom2 = Button(frame_extrac_excel, text="Voltar",
                              height=1, command=lambda: self.switch_frame(0))
        back_buttom2.place(x=20, y=100)

        back_buttom3 = Button(frame_modelo_mensagem, text="Voltar",
                              height=1, command=lambda: self.switch_frame(0))
        back_buttom3.place(x=20, y=100)

        back_buttom4 = Button(frame_about, text="Voltar",
                              height=1, command=lambda: self.switch_frame(0))
        back_buttom4.place(x=410, y=115, width=80)

        # select messagem 🖼
        label_descri = Label(frame_modelo_mensagem, text=(
            "Selecione o arquivo de mensagem(Ex: .txt, .htm, .html): "), foreground="black", font=('Helvetica 9'))
        label_descri.place(x=20, y=10)

        escolher_buttom = Button(
            frame_modelo_mensagem, text="Escolher", command=self.onOpen, height=1)
        escolher_buttom.place(x=400, y=32, width=80)

        self.mensagem_file_name_field = Text(
            frame_modelo_mensagem, height="1", foreground="black", borderwidth=1, bg="white")
        self.mensagem_file_name_field.place(x=20, y=35, width=380)

        back_buttom1 = Button(
            frame_modelo_mensagem, text="Salvar", height=1, command=self.save_service)
        back_buttom1.place(x=430, y=110)

        # serviço de email frame

        label_descri = Label(frame_servers, text=(
            "Selecione um serviço de email:"), foreground="black", font=('Helvetica 9'))
        label_descri.place(x=20, y=20)

        self.services = ['Gmail', 'Outlook (Windows 8,10)', 'Outro']
        self.variable = StringVar()
        self.variable.set(self.services[0])
        self.op_menu_server = OptionMenu(
            frame_servers, self.variable, *self.services, command=self.onChangeServer)
        self.op_menu_server.place(x=200, y=16, width=200)

        self.label_email_service = Label(frame_servers, text=(
            "Email: "), foreground="black", font=('Helvetica 9'))
        self.service_email = Text(
            frame_servers, height="1", foreground="black", borderwidth=1)
        self.label_password_service = Label(frame_servers, text=(
            "Senha: "), foreground="black", font=('Helvetica 9'))
        self.service_password = Text(
            frame_servers, height="1", foreground="black", borderwidth=1)

        self.label_descri_service = Label(frame_servers, text=(
            "Serviço (Ex: gmail.com): "), foreground="black", font=('Helvetica 9'))
        self.service_name = Text(
            frame_servers, height="1", foreground="black", borderwidth=1)
        self.service_port_label = Label(frame_servers, text=(
            "Porta smtp (Ex: 587): "), foreground="black", font=('Helvetica 9'))
        self.service_port = Text(
            frame_servers, height="1", foreground="black", borderwidth=1)

        self.label_email_service.place(x=20, y=55)
        self.service_email.place(x=80, y=55, width=150)
        self.label_password_service.place(x=250, y=55)
        self.service_password.place(x=300, y=55, width=120)

        back_buttom1 = Button(frame_servers, text="Salvar",
                              height=1, command=self.save_service)
        back_buttom1.place(x=430, y=110)

        # campos de configurações sobre colunas no excel ⁉

        self.label_pagina_excel = Label(frame_extrac_excel, text=(
            "Nome da pagina: "), foreground="black", font=('Helvetica 9'))
        self.pagina_excel = Text(
            frame_extrac_excel, height="1", foreground="black", borderwidth=1)

        self.label_coluna_email_excel = Label(frame_extrac_excel, text=(
            "Coluna dos Emails: "), foreground="black", font=('Helvetica 9'))
        self.coluna_email_excel = Text(
            frame_extrac_excel, height="1", foreground="black", borderwidth=1)

        self.label_coluna_preenchimento = Label(frame_extrac_excel, text=(
            "Colunas de preenchimento: (Ex: A;Z;J): "), foreground="black", font=('Helvetica 9'))
        self.coluna_preenchimento = Text(
            frame_extrac_excel, height="1", foreground="black", borderwidth=1)

        self.label_pagina_excel.place(x=20, y=20)
        self.pagina_excel.place(x=130, y=20, width=90)

        self.label_coluna_email_excel.place(x=240, y=20)
        self.coluna_email_excel.place(x=360, y=20, width=70)

        self.label_coluna_preenchimento.place(x=20, y=50)
        self.coluna_preenchimento.place(x=240, y=50, width=190)

        back_buttom1 = Button(frame_extrac_excel, text="Salvar",
                              height=1, command=self.save_service)
        back_buttom1.place(x=430, y=110)

        # sobre

        label_descri = Label(frame_about, text=("Sobre:"),
                             foreground="black", font=('Helvetica 11 bold'))
        label_descri.place(x=20, y=10)

        label_descri = Label(frame_about, text=(
            "Descrição:"), foreground="black", font=('Helvetica 10 bold'))
        label_descri.place(x=25, y=28)

        label_descri = Label(frame_about, text=("Ferramenta desenvolvida para automatizar o envio de emails oriundos do excel."),
                             foreground="black", font=('Helvetica 9'), justify="left")
        label_descri.place(x=35, y=45,)

        label_descri = Label(frame_about, text=(
            "Desenvolvimento:"), foreground="black", font=('Helvetica 10 bold'))
        label_descri.place(x=25, y=60)

        label_descri = Label(frame_about, text=("Laboratório de Tecnologia Contra lavagem de Dinheiro da Policia Civil do RN."),
                             foreground="black", font=('Helvetica 9'), justify="left")
        label_descri.place(x=35, y=80)

        label_descri = Label(frame_about, text=(
            "Desenvolvedor:"), foreground="black", font=('Helvetica 10 bold'))
        label_descri.place(x=25, y=95)

        label_descri = Label(frame_about, text=("Christian Lima Freitag - christianfreitag2019@gmail.com"),
                             foreground="black", font=('Helvetica 10'), justify="left")
        label_descri.place(x=35, y=115)


def main():

    root = Tk()
    root.geometry("500x150")
    root.configure(bg="#F5F5F5")
    root.resizable(False, False)
    root.title("LabTools - Exmail")
    # root.call('wm', 'iconphoto', root._w,PhotoImage(data=IMAGE_ENCODED))
    app = Exmail(root)
    root.mainloop()


if __name__ == '__main__':
    main()
