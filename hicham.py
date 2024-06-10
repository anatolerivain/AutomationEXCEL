import re
import pdfplumber
from openpyxl import Workbook

# Chemin vers le fichier PDF
input_pdf = '/Users/arivain/Desktop/hicham.pdf'
output_file = 'hicham.xlsx'

# Lire le contenu du fichier PDF
content = ''
with pdfplumber.open(input_pdf) as pdf:
    for page in pdf.pages:
        content += page.extract_text() + '\n'

# Trouver toutes les lignes sous "PLACARD"
pattern = re.compile(r'PLACARD\n(.*)')
matches = pattern.findall(content)
print("matches: ", matches)

# Créer un nouveau classeur Excel
workbook = Workbook()
sheet = workbook.active
sheet.title = "Placard"

# Écrire les lignes extraites dans le fichier Excel
for index, line in enumerate(matches, start=1):
    sheet.cell(row=index, column=1, value=line)

# Enregistrer le fichier Excel
workbook.save(output_file)

print(f"Les lignes ont été extraites et enregistrées dans {output_file}")
