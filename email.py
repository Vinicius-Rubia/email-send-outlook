import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)

email.To = 'email_1;email_2'
email.Subject = 'Envio de email com Outlook'
email.HTMLBody = '''
<p>Olá, Aqui é o Vinicius</p>
<p>É possível usarmos HTML no envio de email</p>
<p>Podemos também enviar anexos</p>
<p>Tudo isso com o python</p>

<p>Abs,</p>
<p>Vinicius rubia</p>
'''
anexo = r'caminho_arquivo_anexo'

email.Attachments.Add(anexo)

email.Send()
print('Email Enviado')