from imap_tools import MailBox, AND
from datetime import date

# Instancia da mailbox autenticada.
mailbox = MailBox('IMAP').login('USERNAME', 'PASSWORD')

# Seleciona a pasta onde será feito o get
mailbox.folder.set('INBOX')

# Utilizando método fetch() com AND para retornar uma lista de emails onde todos argumentos sejam verdadeiros:
# from_ (remetente) == remetente@domain.com
# subject (título) == Hello
# date (data) == date.today() <Data de Hoje>
mails = mailbox.fetch(AND(from_='remetente@domain.com', subject='Hello', date=date.today()))

# Para cada email na lista de email:
for mail in mails:
    # print(mail.subject) -> Escreve o título do email no console
    print(mail.subject)
    # print(mail.text) -> Escreve o corpo do email no console
    print(mail.text)

# Para cada email na lista de email:
for mail in mails:

    # Se o email tiver mais de 0 anexos
    if len(mail.attachments) > 0:

        # Para cada anexo na lista de anexos do email
        for att in mail.attachments:

            # Se o nome do anexo for ArquivoExcel
            if "ArquivoExcel" in att.filename:

                # Armazena na lista atts os bytes de cada anexo
                info_att = att.payload

                # Criar um arquivo.xlsx escrevendo em bytes e nomeando como PlanilhaExcel
                with open('arquivo.xlsx', 'wb') as PlanilhaExcel:
                    # Escrevendo na PlanilhaExcel o conteudo de info_att
                    PlanilhaExcel.write(info_att)
