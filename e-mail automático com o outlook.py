import os
import win32com.client as win32

remetente = input("Digite os remetentes separados por ';' : ")
assunto = input("Digite o assunto de e-mail: ")
mensagem = input("Digite o corpo do e-mail: ")

# faz a integralização com o outlook
outlook = win32.Dispatch('outlook.application')

# cria um e-mail
email = outlook.CreateItem(0)

# Adiciona um remetente ao e-mail
email.To = f"{remetente}"
# Adiciona um assunto ao e-mail
email.Subject = f"{assunto}"
# Configura o corpo do e-mail com HTML
email.HTMLBody = f"""
<p>{mensagem}</p>
<p>Abs,</p>
<p>Código Python</p>
<p>E-mail automático do Python.</p>
"""
# Opção para adicionar anexos ao e-mail
escolha_anexo = int(input("Digite 1 se quiser anexar um arquivo no e-mail: "))
if escolha_anexo == 1:
    numero_de_anexos = int(input("Digite a quantidade de anexos que quer colocar no e-mail: "))
    m = 0
    while m < numero_de_anexos:
        m = m + 1
        anexo = input("Digite o caminho do arquivo para anexar no e-mail: ")    # Pega o endereço dos arquivos de anexo

        caracter = '"'      # Remove as aspas " do endereço do arquivo
        for i in caracter:
            anexo = anexo.replace(i, '')

        email.Attachments.Add(anexo)        # Insere o arquivo em anexo

email.Send()        # Envia o e-mail
print("Email Enviado")
os.system("pause")
