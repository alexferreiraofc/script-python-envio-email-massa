mport smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import time

#email que vai disparar a mensagem
remetente = "naoresponda@mail.com"

def enviar_email():
    try:

        #Pandas
        df = pd.read_excel("emails.xlsx")

        for _, row in df.iterrows():
            email = row["Email"]


            try:
                # Configura a mensagem do e-mail
                    assunto = "Imóveis Férias coletivas 2024"
                    corpo = f"""
                <html>
                    <body>
                        <p>mensagem envio. etc...</p>
                        <img src="cid:imagem" alt="Imagem de férias" style="width:400px;">
                    </body>
                </html>
                """
                    
                    # Define a mensagem com codificação UTF-8
                    msg = MIMEMultipart("related")
                    msg['Subject'] = assunto
                    msg['From'] = remetente
                    msg['To'] = email

                    msg_alternative = MIMEMultipart("alternative")
                    msg.attach(msg_alternative)
                    msg_alternative.attach(MIMEText(corpo, "html"))


                    with open("ferias.jpg", "rb") as imagem:
                        mime_base = MIMEBase("image", "jpeg", filename="ferias.jpg")
                        mime_base.add_header("Content-ID", "<imagem>")  # ID usado no HTML
                        mime_base.add_header("Content-Disposition", "inline", filename="ferias.jpg")
                        mime_base.set_payload(imagem.read())
                        encoders.encode_base64(mime_base)
                        msg.attach(mime_base)


                    # Configura o servidor SMTP
                    # ATENÇÃO ao limite de envios por hora!!!
                    with smtplib.SMTP("smtp.servidor.com.br", 587) as servidor:
                        servidor.starttls()  # Habilita criptografia STARTTLS
                        servidor.login("naoresponda@mail.com.br", "senhadoemail")
                        servidor.sendmail(remetente, email, msg.as_string())

                    print(f"E-mail enviado com sucesso para {email}")

                    time.sleep(2)
            except Exception as e:
                    print(f"Erro ao enviar e-mail: {email}: {e}")

    except Exception as e:
        print(f"Erro ao processar a planilha: {e}")


enviar_email()