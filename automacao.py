import win32com.client as win32

# criando integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um e-mail
email = outlook.CreateItem(0)

# configurar informações do e-mail
email.To = "isasilva123ptc@gmail.com"
email.Subject = "E-mail teste Phyton"
email.HTMLBody = """"
<p>Ola Isa</p>
<p>Teste Phyton</p>
"""
email.Send()
print("E-mail enviado com sucesso")
