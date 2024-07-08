from flask import render_template, request, send_file
from app import app
from app.email_extractor import EmailExtractor
from app.attachment_extractor import AttachmentExtractor
from app.excel_exporter import ExcelExporter
from config import USUARIO, SENHA
from imap_tools import MailBox, AND
from io import BytesIO
import logging
import re

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/extract', methods=['POST'])
def extract_data():
    email = request.form['email']

    if not is_valid_email(email):
        return {"error": "Por favor, insira um e-mail válido."}, 400

    try:
        # Faça login na caixa de correio
        with MailBox('imap.gmail.com').login(USUARIO, SENHA) as mailbox:
            # Busque e-mails do remetente especificado
            emails = mailbox.fetch(AND(from_=email, text="email@news-voeazul.com.br"), limit=100)

            extracted_data = []

            for email_msg in emails:
                # Extrair informações do corpo do e-mail
                email_extractor = EmailExtractor(email_msg.text)
                info = email_extractor.extract_info()
                if info:
                    extracted_data.append(info)

                for attachment in email_msg.attachments:
                    # Verificar se o anexo é um arquivo .eml
                    if attachment.filename.lower().endswith('.eml'):
                        attachment_extractor = AttachmentExtractor(attachment.payload)
                        info = attachment_extractor.extract_info_from_attachment()
                        if info:
                            extracted_data.append(info)

            # Verificar se a lista de dados extraídos não está vazia antes de exportar
            if extracted_data:
                excel_file = BytesIO()
                excel_exporter = ExcelExporter(extracted_data)
                excel_exporter.export_to_excel(excel_file)
                excel_file.seek(0)
                return send_file(excel_file, download_name='dados_extraidos.xlsx', as_attachment=True)
            else:
                return {"error": "Nenhum dado encontrado para exportar."}, 404

    except Exception as e:
        logging.error(f"Erro ao fazer login ou buscar e-mails: {e}")
        return {"error": "Ocorreu um erro ao processar a solicitação. Por favor, tente novamente mais tarde."}, 500

def is_valid_email(email):
    email_regex = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
    return re.match(email_regex, email)