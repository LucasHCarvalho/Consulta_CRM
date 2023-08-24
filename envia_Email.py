import os
import win32com.client
from datetime import datetime

# Importe a classe Outlook Application
outlook = win32com.client.Dispatch("Outlook.Application")

# Crie um novo e-mail
mail = outlook.CreateItem(0)

# Anexe o arquivo CSV
anexo = os.path.abspath('crm.csv')
mail.Attachments.Add(anexo)

# Assunto do e-mail
mail.Subject = "Crm Inativo Cremesp"

# Parametros do Corpo do Email
data_e_hora_atuais = datetime.now()
data_e_hora_atuais = data_e_hora_atuais.strftime("%H:%M")

if data_e_hora_atuais < '12:00':
    saudacao = 'Bom dia!'
if data_e_hora_atuais < '18:00':
    saudacao = 'Boa tarde!'
else:
    saudacao = 'Boa noite!'

texto = ''' 

Identificamos através de uma consulta no Cremesp que os médicos da relação, encontram-se com o seu CRM inativo no Conselho Regional de Medicina.
Os mesmos estão cadastrados e com situação Ativa dentro da plataforma SOUL-MV.
'''
assinatura = '''
'''

# Corpo do e-mail
mail.Body = saudacao + texto + assinatura

# Defina o destinatário do e-mail
mail.To = "email.sobrenome@mail.com.br"

# Envie o e-mail
mail.Send()
