import sys
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, 
                             QTableWidget, QTableWidgetItem, QMessageBox, QHBoxLayout, QMenu, QAction, QInputDialog)
from PyQt5.QtCore import Qt

class ExcelViewerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.df1 = None
        self.df2 = None
        self.copied_data = []
        self.current_table = None

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Visor de Excel Mejorado')
        self.setGeometry(100, 100, 1200, 600)

        layout = QVBoxLayout()

        button_layout = QHBoxLayout()
        
        self.load_button_1 = QPushButton('Adjuntar Archivo EXCEL - 1', self)
        self.load_button_1.clicked.connect(lambda: self.load_file(1))
        button_layout.addWidget(self.load_button_1)

        self.load_button_2 = QPushButton('Adjuntar Archivo EXCEL - 2', self)
        self.load_button_2.clicked.connect(lambda: self.load_file(2))
        button_layout.addWidget(self.load_button_2)

        self.copy_button = QPushButton('Copiar Selección', self)
        self.copy_button.clicked.connect(self.copy_selection)
        button_layout.addWidget(self.copy_button)

        self.paste_button = QPushButton('Pegar en Excel 2', self)
        self.paste_button.clicked.connect(self.paste_selection)
        button_layout.addWidget(self.paste_button)

        self.add_row_button = QPushButton('Añadir Fila a Excel 2', self)
        self.add_row_button.clicked.connect(lambda: self.add_row_to_table(self.table2, self.df2))
        button_layout.addWidget(self.add_row_button)

        self.save_button = QPushButton('Guardar Archivo EXCEL - 2', self)
        self.save_button.clicked.connect(self.save_file)
        button_layout.addWidget(self.save_button)

        layout.addLayout(button_layout)

        self.table1 = QTableWidget(self)
        self.table1.setDragEnabled(True)
        self.table1.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table1.customContextMenuRequested.connect(lambda pos: self.show_context_menu(pos, self.table1))
        layout.addWidget(self.table1)

        self.table2 = QTableWidget(self)
        self.table2.setAcceptDrops(True)
        self.table2.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table2.customContextMenuRequested.connect(lambda pos: self.show_context_menu(pos, self.table2))
        layout.addWidget(self.table2)

        self.setLayout(layout)

    def load_file(self, file_number):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Abrir Archivo Excel", "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            if file_number == 1:
                self.df1 = df
                self.display_dataframe(df, self.table1)
                QMessageBox.information(self, "Éxito", "Archivo 1 cargado correctamente")
            elif file_number == 2:
                self.df2 = df
                self.display_dataframe(df, self.table2)
                QMessageBox.information(self, "Éxito", "Archivo 2 cargado correctamente")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cargar el archivo: {e}")

    def display_dataframe(self, df, table):
        table.setRowCount(df.shape[0])
        table.setColumnCount(df.shape[1])
        table.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

    def copy_selection(self):
        selected_items = self.table1.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Advertencia", "Debe seleccionar al menos una celda.")
            return

        self.copied_data = [(item.row(), item.column(), item.text()) for item in selected_items]
        QMessageBox.information(self, "Éxito", "Celdas copiadas correctamente")

    def paste_selection(self):
        if not self.copied_data or self.df2 is None:
            QMessageBox.critical(self, "Error", "Debe copiar celdas del primer archivo y cargar el segundo archivo.")
            return

        try:
            for row, col, value in self.copied_data:
                self.df2.iat[row, col] = value

            self.display_dataframe(self.df2, self.table2)
            QMessageBox.information(self, "Éxito", "Datos pegados correctamente.")
            self.copied_data = []
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo pegar el dato: {e}")

    def add_row_to_table(self, table, df):
        if df is None:
            QMessageBox.critical(self, "Error", "Debe cargar el archivo primero.")
            return

        new_row = pd.DataFrame([[""] * df.shape[1]], columns=df.columns)
        df = pd.concat([df, new_row], ignore_index=True)
        self.display_dataframe(df, table)
        if table == self.table1:
            self.df1 = df
        else:
            self.df2 = df

    def save_file(self):
        self.update_dataframe_from_table(self.table2, self.df2)
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Archivo Excel", "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if not file_path:
            return

        try:
            self.df2.to_excel(file_path, index=False)
            QMessageBox.information(self, "Éxito", "Archivo guardado correctamente")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {e}")

    def update_dataframe_from_table(self, table, df):
        if df is None:
            return

        # Verificar si el DataFrame tiene las mismas dimensiones que la tabla
        if table.rowCount() != df.shape[0] or table.columnCount() != df.shape[1]:
            # Ajustar el DataFrame a las dimensiones de la tabla
            new_df = pd.DataFrame(index=range(table.rowCount()), columns=df.columns[:table.columnCount()])
        else:
            new_df = df

        for i in range(table.rowCount()):
            for j in range(table.columnCount()):
                item = table.item(i, j)
                new_df.iat[i, j] = item.text() if item else ""

        # Actualizar el DataFrame correspondiente
        if table == self.table1:
            self.df1 = new_df
        else:
            self.df2 = new_df

    def show_context_menu(self, position, table):
        context_menu = QMenu(self)

        add_row_action = QAction('Añadir Fila', self)
        add_row_action.triggered.connect(lambda: self.add_row_to_table(table, self.get_df(table)))
        context_menu.addAction(add_row_action)

        add_column_action = QAction('Añadir Columna', self)
        add_column_action.triggered.connect(lambda: self.add_column_to_table(table))
        context_menu.addAction(add_column_action)

        edit_cell_action = QAction('Editar Celda', self)
        edit_cell_action.triggered.connect(lambda: self.edit_cell(table))
        context_menu.addAction(edit_cell_action)

        delete_row_action = QAction('Eliminar Fila', self)
        delete_row_action.triggered.connect(lambda: self.delete_row_from_table(table))
        context_menu.addAction(delete_row_action)

        delete_column_action = QAction('Eliminar Columna', self)
        delete_column_action.triggered.connect(lambda: self.delete_column_from_table(table))
        context_menu.addAction(delete_column_action)

        context_menu.exec_(table.viewport().mapToGlobal(position))

    def get_df(self, table):
        if table == self.table1:
            return self.df1
        elif table == self.table2:
            return self.df2
        return None

    def add_column_to_table(self, table):
        df = self.get_df(table)
        if df is None:
            QMessageBox.critical(self, "Error", "Debe cargar el archivo primero.")
            return

        column_name, ok = QInputDialog.getText(self, 'Añadir Columna', 'Nombre de la nueva columna:')
        if ok and column_name:
            df[column_name] = ""
            self.display_dataframe(df, table)
            if table == self.table1:
                self.df1 = df
            else:
                self.df2 = df

    def edit_cell(self, table):
        row = table.currentRow()
        col = table.currentColumn()
        if row < 0 or col < 0:
            QMessageBox.warning(self, "Advertencia", "Seleccione una celda para editar.")
            return

        current_value = table.item(row, col).text()
        new_value, ok = QInputDialog.getText(self, 'Editar Celda', 'Nuevo valor:', text=current_value)
        if ok:
            table.setItem(row, col, QTableWidgetItem(new_value))
            df = self.get_df(table)
            df.iat[row, col] = new_value

    def delete_row_from_table(self, table):
        row = table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Advertencia", "Seleccione una fila para eliminar.")
            return

        table.removeRow(row)
        df = self.get_df(table)
        df.drop(df.index[row], inplace=True)
        df.reset_index(drop=True, inplace=True)
        self.display_dataframe(df, table)

    def delete_column_from_table(self, table):
        col = table.currentColumn()
        if col < 0:
            QMessageBox.warning(self, "Advertencia", "Seleccione una columna para eliminar.")
            return

        df = self.get_df(table)
        df.drop(df.columns[col], axis=1, inplace=True)
        self.display_dataframe(df, table)

    def dragEnterEvent(self, event):
        event.accept()

    def dragMoveEvent(self, event):
        event.accept()

    def dropEvent(self, event):
        if event.source() == self.table1:
            selected_items = self.table1.selectedItems()
            if not selected_items:
                return

            for item in selected_items:
                row, col = item.row(), item.column()
                text = item.text()
                self.df2.iat[row, col] = text

            self.display_dataframe(self.df2, self.table2)
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = ExcelViewerApp()
    viewer.show()
    sys.exit(app.exec_())
