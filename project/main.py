import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Crear una ventana de diálogo para que el usuario seleccione el archivo
root = Tk()
root.withdraw()
filename = askopenfilename()

# Leer el archivo seleccionado con Pandas
df = pd.read_excel(filename)

# Continuar con el resto del código tal y como lo tenías antes
df2 = df[['Delivery', 'External HU']].dropna()
df3 = df2.groupby('Delivery')['External HU'].nunique().reset_index(name='External HU Unicos')

# Guardar el resultado en una nueva hoja en el mismo archivo de Excel
with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
    df3.to_excel(writer, sheet_name='Unique HU', index=False)

print('Se ejecutó correctamente')
print('Buen día!')