import os
import win32com.client as win32

remetente = input("Digite os remetentes separados por ';' : ")
assunto = input("Digite o assunto de e-mail: ")
mensagem = input("Digite o corpo do e-mail: ")

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

email.To = f"{remetente}"
email.Subject = f"{assunto}"
email.HTMLBody = f"""
<p>{mensagem}</p>
<p>Abs,</p>
<p>Código Python</p>
<p>E-mail automático do Python.</p>
"""
escolha_anexo = int(input("Digite 1 se quiser anexar um arquivo no e-mail: "))
if escolha_anexo == 1:
    numero_de_anexos = int(input("Digite a quantidade de anexos que quer colocar no e-mail: "))
    m = 0
    while m < numero_de_anexos:
        m = m + 1
        anexo = input("Digite o caminho do arquivo para anexar no e-mail: ")

        caracter = '"'
        for i in caracter:
            anexo = anexo.replace(i, '')

        email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")
os.system("pause")
