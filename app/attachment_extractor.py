from email import policy, parser
from bs4 import BeautifulSoup
import re
import logging

class AttachmentExtractor:
    def __init__(self, attachment):
        self.attachment = attachment

    def extract_text_from_eml(self):
        try:
            msg = parser.BytesParser(policy=policy.default).parsebytes(self.attachment)

            # Verifica se é multipart para iterar nas partes
            if msg.is_multipart():
                for part in msg.iter_parts():
                    # Verifica se é texto e procura o campo de subject
                    if part.get_content_type() in ('text/plain', 'text/html'):
                        payload = part.get_payload(decode=True)
                        # Aqui você pode verificar a presença do campo de subject
                        # Por exemplo, se estiver em HTML, procure por "<title>" ou "<subject>"
                        # Ou em text/plain por "Subject:"
                        return payload.decode('utf-8', errors='ignore')
            else:
                return msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        except AttributeError as e:
            logging.error(f"Erro ao extrair texto do anexo (AttributeError): {e}")
            return None
        except LookupError as e:
            logging.error(f"Erro ao extrair texto do anexo (LookupError): {e}")
            return None
        except Exception as e:
            logging.error(f"Erro ao extrair texto do anexo: {e}")
            return None

    def extract_values(self, html):
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
        except AttributeError as e:
            logging.error(f"Erro ao extrair valores do anexo (AttributeError): {e}")
            return None, None, None, None, None, None
        except TypeError as e:
            logging.error(f"Erro ao extrair valores do anexo (TypeError): {e}")
            return None, None, None, None, None, None
        except Exception as e:
            logging.error(f"Erro ao extrair valores do anexo: {e}")
            return None, None, None, None, None, None

    def extract_info_from_attachment(self):
        attachment_text = self.extract_text_from_eml()
        if attachment_text:
            tarifa_total, taxas, localizador, origem, destino, passageiros = self.extract_values(attachment_text)

            return {
                'Localizador': localizador,
                'Origem': origem,
                'Destino': destino,
                'Passageiros': ", ".join(passageiros),
                'Milhas': tarifa_total,
                'Taxas': taxas
            }
        else:
            return None