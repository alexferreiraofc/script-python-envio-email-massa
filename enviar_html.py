import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# Define um timer entre cada envio
import time


# dados da conta de e-mail
sender_email = "seuemail@mail.com.br"
password = "senha"

# carregando lista de e-mails de um arquivo excel
df = pd.read_excel("planilha.xlsx")
emails = df["email"].tolist()

# Caso eu queira cópia para algum email
cc_email = ""

# configurações do servidor Zimbra
server = smtplib.SMTP_SSL("servidor.smtp", 000)
server.login(sender_email, password)

# Faz a leitura da imagem da assinatura
with open("assinatura.png", 'rb') as f:
    img_data = f.read()
image = MIMEImage(img_data, name="signature")

# loop para enviar e-mails para cada endereço na lista em HTML
for email in emails:

    html = """\
    <html>
      <body>
        <p>Prezado Cliente,<br><br>
           Esse email é apenas um modelo, haja vista que não posso disponibilizar informações da empresa
           </p>
      </body>
    </html>
    """
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = email
    message['Cc'] = cc_email
    # Assunto do email
    message['Subject'] = "IPTU - Cambirela Imóveis"
    message.attach(MIMEText(html, 'html'))
    message.attach(image)
    server.sendmail(sender_email, [email, cc_email], message.as_string())

    # confirmação de envio + email destino
    print("Email enviado para... [", email, "]... Reiniciando Loop...")

    # timer para evitar blacklist
    time.sleep(5)


server.quit()

# Confirma finalização de todo o processo de envio
print("E-mails enviados com sucesso!")
