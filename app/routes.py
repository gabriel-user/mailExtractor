import logging
import re
from io import BytesIO
from flask import render_template, request, send_file
from imap_tools import MailBox, AND

from app import app
from app.email_extractor import EmailExtractor
from app.attachment_extractor import AttachmentExtractor
from app.excel_exporter import ExcelExporter

USUARIO = 'conciliacaoncm@maxmilhas.com.br'
SENHA = "vhlm elbm fugg mddy"


@app.route('/', methods=['GET'])
def index():
    """Rota principal da aplicação."""
    return render_template('index.html')


@app.route('/extract', methods=['POST'])
def extract_data():
    """Rota para extrair dados de e-mails."""
    email = request.form['email']

    if not is_valid_email(email):
        return {"error": "Por favor, insira um e-mail válido."}, 400

    try:
        # Faça login na caixa de correio
        with MailBox('imap.gmail.com').login(USUARIO, SENHA) as mailbox:
            # Busque e-mails do remetente especificado
            emails = mailbox.fetch(AND(from_=email), limit=100)

            extracted_data = []
            localizadores_existentes = {}

            for email_msg in emails:
                for attachment in email_msg.attachments:
                    # Verificar se o anexo é um arquivo .eml
                    if attachment.filename.lower().endswith('.eml'):
                        attachment_extractor = AttachmentExtractor(attachment.payload)
                        attachment_info = attachment_extractor.extract_info_from_attachment()
                        if attachment_info:
                            localizador = attachment_info['Localizador']
                            tipo_movimentacao = attachment_info['Tipo de movimentação']
                            if localizador not in localizadores_existentes or (localizador in localizadores_existentes and localizadores_existentes[localizador] != tipo_movimentacao):
                                extracted_data.append(attachment_info)
                                localizadores_existentes[localizador] = tipo_movimentacao

                # Extrair informações do corpo do e-mail
                email_extractor = EmailExtractor(email_msg.text)
                email_info = email_extractor.extract_info()
                if email_info:
                    localizador = email_info['Localizador']
                    tipo_movimentacao = email_info['Tipo de movimentação']
                    if localizador not in localizadores_existentes or (localizador in localizadores_existentes and tipo_movimentacao == 'Cancelamento'):
                        extracted_data.append(email_info)
                        localizadores_existentes[localizador] = tipo_movimentacao

            # Verificar se a lista de dados extraídos não está vazia antes de exportar
            if extracted_data:
                excel_file = BytesIO()
                excel_exporter = ExcelExporter(extracted_data)
                excel_exporter.export_to_excel(excel_file)
                excel_file.seek(0)
                return send_file(excel_file, download_name='dados_extraidos.xlsx', as_attachment=True)
            else:
                return {"error": "Nenhum dado encontrado para exportar."}, 404

    except Exception as error:
        logging.error(f"Erro ao fazer login ou buscar e-mails: {error}")
        return {
            "error": "Ocorreu um erro ao processar a solicitação. Por favor, tente novamente mais tarde."
        }, 500


def is_valid_email(email):
    """Verifica se um e-mail é válido."""
    email_regex = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
    return re.match(email_regex, email)


@app.errorhandler(404)
def page_not_found(error):
    return render_template('404.html'), 404


app.logger.setLevel(logging.ERROR)


@app.errorhandler(500)
def internal_server_error(error):
    app.logger.error(f"Internal Server Error: {error}")
    return render_template('500.html'), 500