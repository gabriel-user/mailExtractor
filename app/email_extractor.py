import re
import logging


class EmailExtractor:
    """Classe para extrair informações de um e-mail."""

    def __init__(self, email_text):
        self.email_text = email_text

    def extract_date(self):
        """Extrai a data do e-mail."""
        date_regex = r'Date:\s*(.+)'
        match = re.search(date_regex, self.email_text)
        if match:
            return match.group(1)
        return None

    def extract_info(self):
        """Extrai informações do texto do e-mail."""
        # Expressões regulares para capturar as informações
        locator_regex = r'Reserva\s+([A-Z0-9]+)\s+Realizada com Sucesso'
        origin_regex = r'Origem:\s*(\w+)'
        destination_regex = r'Destino:\s*(\w+)'
        passengers_regex = r'\b([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)\b'
        miles_regex = r'Tarifa Total\s*(\d{1,3}(?:\.\d{3})*) pontos'
        taxes_regex = r'Taxas\s*R\$ (\d{1,3}(?:\.\d{3})*,\d{2})'

        try:
            # Encontrar as informações no texto do e-mail
            locator = re.search(locator_regex, self.email_text)
            origin = re.search(origin_regex, self.email_text)
            destination = re.search(destination_regex, self.email_text)
            passengers = re.findall(passengers_regex, self.email_text)
            miles = re.search(miles_regex, self.email_text)
            taxes = re.search(taxes_regex, self.email_text)
            date = self.extract_date()

            # Formatar os nomes dos passageiros
            formatted_passengers = [name.strip() for name in passengers]

            # Concatenar todos passageiros
            concatenated_passengers = ", ".join(formatted_passengers)

            # Extrair os grupos encontrados ou None se não encontrado
            locator = locator.group(1) if locator else None
            origin = origin.group(1) if origin else None
            destination = destination.group(1) if destination else None
            miles = miles.group(1) if miles else None
            taxes = taxes.group(1) if taxes else None

            return {
                'Localizador': locator,
                'Origem': origin,
                'Destino': destination,
                'Passageiros': concatenated_passengers,
                'Milhas': miles,
                'Taxas': taxes,
                'Data': date
            }
        except Exception as error:
            logging.error(f"Erro ao extrair informações do e-mail: {error}")
            raise