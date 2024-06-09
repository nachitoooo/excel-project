import pandas as pd
import openpyxl

def read_excel_with_styles(file_path):
    # Cargar el archivo Excel con openpyxl
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    # Leer los nombres de las columnas
    column_names = []
    for cell in sheet[1]:  # Asumiendo que la primera fila contiene los nombres de las columnas
        column_names.append(cell.value)

    # Leer los datos en un DataFrame
    df = pd.DataFrame(sheet.iter_rows(min_row=2, values_only=True), columns=column_names)
    
    return df, column_names

# Cargar el archivo Excel
file_path = 'test1.xlsx'  # Actualiza con la ruta correcta
df, column_names = read_excel_with_styles(file_path)

# Mostrar los nombres de las columnas
print("Columnas del archivo Excel:")
for col in column_names:
    print(col)

# Mostrar los datos
print("\nDatos del archivo Excel:")
print(df)
