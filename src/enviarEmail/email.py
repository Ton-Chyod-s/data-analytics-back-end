import smtplib
import email.message

def enviar_email(assunto, de, para, senha,corpo_mensagem):  
    corpo_email = corpo_mensagem
    msg = email.message.Message()
    msg['Subject'] = assunto
    msg['From'] = de
    msg['To'] = para
    password = senha
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

assunto = "Atualização DRE-servidor"
para = "klayton.oliveira@perincontabil.com.br"
de = "perindevboot@gmail.com"
senha = "gxkqsyymnogquthd"