import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import warnings

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Merger")

        # Variables para almacenar los nombres de los archivos y columnas
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.column1_name = tk.StringVar()
        self.column2_name = tk.StringVar()
        self.new_column_name = tk.StringVar()

        # Crear widgets
        tk.Label(root, text="Archivo 1:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(root, textvariable=self.file1_path, width=50).grid(row=0, column=1)
        tk.Button(root, text="Seleccionar", command=self.select_file1).grid(row=0, column=2)

        tk.Label(root, text="Columna 1:").grid(row=1, column=0, padx=10, pady=10)
        self.column1_menu = tk.OptionMenu(root, self.column1_name, "")
        self.column1_menu.grid(row=1, column=1)

        tk.Label(root, text="Archivo 2:").grid(row=2, column=0, padx=10, pady=10)
        tk.Entry(root, textvariable=self.file2_path, width=50).grid(row=2, column=1)
        tk.Button(root, text="Seleccionar", command=self.select_file2).grid(row=2, column=2)

        tk.Label(root, text="Columna 2:").grid(row=3, column=0, padx=10, pady=10)
        self.column2_menu = tk.OptionMenu(root, self.column2_name, "")
        self.column2_menu.grid(row=3, column=1)

        tk.Label(root, text="Nombre Nueva Columna:").grid(row=4, column=0, padx=10, pady=10)
        tk.Entry(root, textvariable=self.new_column_name, width=50).grid(row=4, column=1)

        tk.Button(root, text="Combinar y Guardar", command=self.combine_and_save).grid(row=5, column=0, columnspan=3, pady=20)

    def read_excel(self, file_path):
        try:
            # Leer todas las hojas del archivo Excel
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            dataframes = []
            for sheet_name in xls.sheet_names:
                with warnings.catch_warnings(record=True) as w:
                    warnings.simplefilter("always")
                    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, engine='openpyxl')  # Leer todo como string para evitar problemas de tipo
                    df.columns = [f'Column_{i}' if 'Unnamed' in col else col for i, col in enumerate(df.columns)]
                    for warning in w:
                        messagebox.showwarning("Advertencia de Excel", f"Advertencia al leer {file_path}, hoja {sheet_name}: {warning.message}")
                dataframes.append(df)
            return pd.concat(dataframes, ignore_index=True)  # Combinar todas las hojas en un DataFrame
        except Exception as e:
            messagebox.showerror("Error", f"Error al leer el archivo {file_path}: {str(e)}")
            return pd.DataFrame()

    def select_file1(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file1_path.set(file_path)
            df = self.read_excel(file_path)
            valid_columns = [col for col in df.columns if col and not col.startswith('Unnamed')]
            print(f"Columnas en {file_path}: {valid_columns}")
            self.update_columns_menu(valid_columns, self.column1_menu, self.column1_name)

    def select_file2(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file2_path.set(file_path)
            df = self.read_excel(file_path)
            valid_columns = [col for col in df.columns if col and not col.startswith('Unnamed')]
            print(f"Columnas en {file_path}: {valid_columns}")
            self.update_columns_menu(valid_columns, self.column2_menu, self.column2_name)

    def update_columns_menu(self, columns, menu, variable):
        menu['menu'].delete(0, 'end')
        for col in columns:
            menu['menu'].add_command(label=col, command=tk._setit(variable, col))
        if columns:
            variable.set(columns[0])

    def combine_and_save(self):
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        col1 = self.column1_name.get()
        col2 = self.column2_name.get()
        new_col_name = self.new_column_name.get()

        if not file1 or not file2 or not col1 or not col2 or not new_col_name:
            messagebox.showerror("Error", "Todos los campos son obligatorios")
            return

        try:
            df1 = self.read_excel(file1)
            df2 = self.read_excel(file2)

            if col1 not in df1.columns or col2 not in df2.columns:
                messagebox.showerror("Error", "Una de las columnas no existe en los archivos seleccionados")
                return

            combined_data = pd.concat([df1[col1], df2[col2]], axis=0, ignore_index=True)
            combined_data = pd.DataFrame(combined_data, columns=[new_col_name])

            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
            if output_file:
                combined_data.to_excel(output_file, index=False)
                messagebox.showinfo("Éxito", f"Archivo guardado como {output_file}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
