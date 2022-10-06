import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def enviar_email(relatorio, destinatario):
    corpo_email = f'''
        <p> Prezados, <p>
        <p> Segue Relatório de Liberação Do Dia: <p>
        <p> {relatorio}
    '''
    msg = email.message.Message()
    msg['Subject'] = 'Relatório Diário de Horário de Liberação'
    msg['from'] = 'ger1@oliveiratrust.com.br'
    msg['to'] = destinatario
    password = ''
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)
    s = smtplib.SMTP('smtp.gmail.com:587')
    s.starttls()
    '''Credenciais de login para enviar o email'''
    s.login(msg['from'], password)
    s.sendmail(msg['from'], [msg['to']], msg.as_string().encode('utf-8'))
    print('Email Enviado')


def enviar_email_com_anexo(anexo, destinatario, aux1, aux2, aux3, aux4, aux5, aux6, aux7, aux8, aux9):
    username = "luiz.araujo@oliveiratrust.com.br"
    password = "Leco1212!"
    mail_from = "ger1.fundos@oliveiratrust.com.br"
    mail_to = destinatario
    mail_subject = f"Relatório Diário de Horário de Liberação {date.today().strftime('%d/%m/%y')}"
    mail_body = "Prezados, Boa Noite" \
                "\n" \
                f"\nSegue Resumo de Liberação Do Dia {date.today().strftime('%d/%m/%y')}:" \
                "\n" \
                f"\n        • Percentual do dia de ATRASADOS: {aux1}%" \
                f"\n        • Percentual do dia de NÃO LIBERADOS: {aux2}% " \
                f"\n        • Percentual do dia de LIBERADOS NO HORÁRIO: {aux3}%" \
                "\n" \
                f"\nTrês fundos com maiores atrasos, em ordem de atraso:" \
                "\n" \
                f"\n        • {aux4} com {aux5} horas de atraso" \
                "\n" \
                f"\n        • {aux6} com {aux7} horas de atraso" \
                "\n" \
                f"\n        • {aux8} com {aux9} horas de atraso" \
                "\n" \
                "\nLembrando que temos uma planilha de acompanhamento diário e de histórico em:" \
                "\n" \
                "\n" \
                r"          \\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Liberação_Maps\ACOMPANHAMENTO.xlsm" \
                "\n" \
                "\n" \
                "\n" \
                "\nAtenciosamente," \
                "\n" \
                "\nLuiz Araujo"


    mail_attachment = anexo
    mail_attachment_name = "Relatorio Diario.xlsx"
    mimemsg = MIMEMultipart()
    mimemsg['From'] = mail_from
    mimemsg['To'] = mail_to
    mimemsg['Subject'] = mail_subject
    mimemsg.attach(MIMEText(mail_body, 'plain'))
    with open(mail_attachment, "rb") as attachment:
        mimefile = MIMEBase('application', 'octet-stream')
        mimefile.set_payload((attachment).read())
        encoders.encode_base64(mimefile)
        mimefile.add_header('Content-Disposition', f"attachment; filename= {mail_attachment_name}")
        mimemsg.attach(mimefile)
        connection = smtplib.SMTP(host='smtp.gmail.com', port=587)
        connection.starttls()
        connection.login(username, password)
        connection.send_message(mimemsg)
        connection.quit()
    print('email enviado')