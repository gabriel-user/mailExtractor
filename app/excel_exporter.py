from openpyxl import Workbook

class ExcelExporter:
    def __init__(self, data):
        self.data = data

    def export_to_excel(self, output):
        workbook = Workbook()
        sheet = workbook.active

        # Escrever os cabeçalhos na primeira linha
        headers = ['Localizador', 'Origem', 'Destino', 'Passageiros', 'Milhas', 'Taxas']
        sheet.append(headers)

        # Escrever os dados nas linhas subsequentes
        for item in self.data:
            row = [item['Localizador'], item['Origem'], item['Destino'], item['Passageiros'], item['Milhas'], item['Taxas']]
            sheet.append(row)

        workbook.save(output)