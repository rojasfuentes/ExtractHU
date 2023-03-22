import openpyxl

# Abre el archivo de Excel
workbook = openpyxl.load_workbook('C:/Users/JFROJAS/Downloads/Master de Recepciones ADC FY23 Q2 (21)')

# Obtén la hoja de trabajo que contiene las celdas problemáticas
worksheet = workbook['DB Bruta']

# Recorre todas las celdas en la hoja de trabajo
for row in worksheet.rows:
    for cell in row:
        # Verifica si la celda está marcada como fecha y su valor está fuera del rango permitido
        if cell.is_date and (cell.value < 0 or cell.value > 2958465):
            # Corrige el valor de la celda
            cell.value = None

# Guarda los cambios en el archivo de Excel
workbook.save('C:/Users/JFROJAS/Downloads/Master de Recepciones ADC FY23 Q2 (21)')
