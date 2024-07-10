import re
import logging
from email import policy, parser
from bs4 import BeautifulSoup
from datetime import datetime


class AttachmentExtractor:
    """Classe para extrair informações de um anexo de e-mail."""

    def __init__(self, attachment):
        self.attachment = attachment

    def extract_text_from_eml(self):
        """Extrai o texto do anexo de e-mail."""
        try:
            msg = parser.BytesParser(policy=policy.default).parsebytes(self.attachment)
            return msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        except Exception as error:
            logging.error(f"Erro ao extrair texto do anexo: {error}")
            return None

    def extract_values(self, html):
        """Extrai os valores do HTML do anexo."""
        try:
            soup = BeautifulSoup(html, 'html.parser')

            tarifa_total = ""
            taxas = ""
            localizador = ""
            origem = ""
            destino = ""
            passageiros = []

            tables = soup.find_all('table')
            for table in tables:
                rows = table.find_all('tr')
                for row in rows:
                    cells = row.find_all('td')
                    if len(cells) == 2:
                        label = cells[0].get_text(strip=True)
                        value = cells[1].get_text(strip=True)

                        if label == "Tarifa Total":
                            tarifa_total = value.replace("pontos", "").strip()
                        elif label == "Taxas":
                            taxas = value.replace("R$", "").replace("\xa0", "").replace(",", ".").strip()
                        elif label == "Origem":
                            origem = value.strip()
                        elif label == "Destino":
                            destino = value.strip()
                        elif label == "Passageiros":
                            passageiros = [p.strip() for p in value.split(",")]

            # Usar expressão regular para extrair o localizador
            localizador_regex = r'<a\s+[^>]*href="[^"]*pnr=([^&"]+)[^>]*>([^<]*)</a>'
            match = re.search(localizador_regex, html)
            if match:
                localizador = match.group(1)

            return tarifa_total, taxas, localizador, origem, destino, passageiros
        except Exception as error:
            logging.error(f"Erro ao extrair valores do anexo: {error}")
            return None, None, None, None, None, None

    def extract_info_from_attachment(self):
        """Extrai as informações do anexo de e-mail."""
        attachment_text = self.extract_text_from_eml()
        if attachment_text:
            tarifa_total, taxas, localizador, origem, destino, passageiros = self.extract_values(attachment_text)

            # Verificar se a frase "cancelada com sucesso" está presente no texto do anexo
            if "cancelada com sucesso" in attachment_text.lower():
                tipo_movimentacao = "Cancelamento"
            else:
                tipo_movimentacao = "Emissão"

            # Extrair a data de recebimento do e-mail
            msg = parser.BytesParser(policy=policy.default).parsebytes(self.attachment)
            data_recebimento = msg['Date']
            data_recebimento = datetime.strptime(data_recebimento, '%a, %d %b %Y %H:%M:%S %z').strftime('%d/%m/%Y')

            return {
                'Data': data_recebimento,
                'Localizador': localizador,
                'Origem': origem,
                'Destino': destino,
                'Passageiros': ", ".join(passageiros),
                'Milhas': tarifa_total,
                'Taxas': taxas,
                'Tipo de movimentação': tipo_movimentacao
            }
        else:
            return None