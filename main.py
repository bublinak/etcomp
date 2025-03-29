import sys
import os
import pandas as pd
import numpy as np
from PyQt5.QtCore import Qt

from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog,
                             QVBoxLayout, QHBoxLayout, QWidget, QLabel, QComboBox,
                             QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView)


class ExcelComparer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Table Comparator")
        self.resize(1000, 800)
        
        self.file1_path = ""
        self.file2_path = ""
        self.df1 = None
        self.df2 = None
        
        self.init_ui()
        
    def init_ui(self):
        main_layout = QVBoxLayout()
        
        # File selection area
        file_layout = QHBoxLayout()
        
        file1_layout = QVBoxLayout()
        self.file1_label = QLabel("První tabulka: Nevybrána")
        self.file1_button = QPushButton("Vyberte první tabulku")
        self.file1_button.clicked.connect(lambda: self.select_file(1))
        file1_layout.addWidget(self.file1_label)
        file1_layout.addWidget(self.file1_button)
        
        file2_layout = QVBoxLayout()
        self.file2_label = QLabel("Druhá tabulka: Nevybrána")
        self.file2_button = QPushButton("Vyberte druhou tabulku")
        self.file2_button.clicked.connect(lambda: self.select_file(2))
        file2_layout.addWidget(self.file2_label)
        file2_layout.addWidget(self.file2_button)
        
        file_layout.addLayout(file1_layout)
        file_layout.addLayout(file2_layout)
        main_layout.addLayout(file_layout)
        
        # Sheet selection area
        sheet_layout = QHBoxLayout()
        
        sheet1_layout = QVBoxLayout()
        sheet1_layout.addWidget(QLabel("Vybraný list z první tabulky:"))
        self.sheet1_combo = QComboBox()
        self.sheet1_combo.currentIndexChanged.connect(self.load_sheet1)
        sheet1_layout.addWidget(self.sheet1_combo)
        
        sheet2_layout = QVBoxLayout()
        sheet2_layout.addWidget(QLabel("Vybraný list z druhé tabulky:"))
        self.sheet2_combo = QComboBox()
        self.sheet2_combo.currentIndexChanged.connect(self.load_sheet2)
        sheet2_layout.addWidget(self.sheet2_combo)
        
        sheet_layout.addLayout(sheet1_layout)
        sheet_layout.addLayout(sheet2_layout)
        main_layout.addLayout(sheet_layout)
        
        # Column selection area
        column_layout = QHBoxLayout()
        
        column1_layout = QVBoxLayout()
        column1_layout.addWidget(QLabel("Prohledávaný sloupec z tabulky 1:"))
        self.column1_combo = QComboBox()
        column1_layout.addWidget(self.column1_combo)
        
        column2_layout = QVBoxLayout()
        column2_layout.addWidget(QLabel("Prohledávaný sloupec z tabulky 2:"))
        self.column2_combo = QComboBox()
        column2_layout.addWidget(self.column2_combo)
        
        column_layout.addLayout(column1_layout)
        column_layout.addLayout(column2_layout)
        main_layout.addLayout(column_layout)
        
        # Matching options
        options_layout = QHBoxLayout()
        self.case_sensitive = QCheckBox("Rozlišovat velká a malá písmena")
        self.numbers_only = QCheckBox("Provnat pouze číselné hodnoty")
        options_layout.addWidget(self.case_sensitive)
        options_layout.addWidget(self.numbers_only)
        main_layout.addLayout(options_layout)
        
        # Compare button
        self.compare_button = QPushButton("Porovnat tabulky")
        self.compare_button.clicked.connect(self.compare_files)
        main_layout.addWidget(self.compare_button)
        
        # Results area
        self.results_label = QLabel("Výsledky porovnání:")
        main_layout.addWidget(self.results_label)
        
        self.results_table = QTableWidget()
        main_layout.addWidget(self.results_table)

        # Export button
        self.export_button = QPushButton("Exportovat výsledky do Excelu")
        self.export_button.clicked.connect(self.export_results)
        main_layout.addWidget(self.export_button)
        
        # Set central widget
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
    
    def select_file(self, file_num):
        file_path, _ = QFileDialog.getOpenFileName(self, f"Vyberte soubor s tabulkami {file_num}", "", "Soubory Excelu (*.xlsx *.xls)")
        if file_path:
            if file_num == 1:
                self.file1_path = file_path
                self.file1_label.setText(f"První tabulka: {os.path.basename(file_path)}")
                self.load_sheets(file_path, self.sheet1_combo)
            else:
                self.file2_path = file_path
                self.file2_label.setText(f"Druhá tabulka: {os.path.basename(file_path)}")
                self.load_sheets(file_path, self.sheet2_combo)
    
    def load_sheets(self, file_path, combo):
        xl = pd.ExcelFile(file_path)
        combo.clear()
        combo.addItems(xl.sheet_names)
    
    def load_sheet1(self):
        if self.file1_path and self.sheet1_combo.currentText():
            self.df1 = pd.read_excel(self.file1_path, sheet_name=self.sheet1_combo.currentText())
            self.column1_combo.clear()
            self.column1_combo.addItems(self.df1.columns)
    
    def load_sheet2(self):
        if self.file2_path and self.sheet2_combo.currentText():
            self.df2 = pd.read_excel(self.file2_path, sheet_name=self.sheet2_combo.currentText())
            self.column2_combo.clear()
            self.column2_combo.addItems(self.df2.columns)
    
    def compare_files(self):
        if self.df1 is None or self.df2 is None:
            self.results_label.setText("Prosím načtěte obě tabulky")
            return
        
        if self.column1_combo.currentText() == "" or self.column2_combo.currentText() == "":
            self.results_label.setText("Prosím vyberte sloupce pro porovnání")
            return
        
        col1 = self.column1_combo.currentText()
        col2 = self.column2_combo.currentText()
        
        # Create copies to avoid modifying originals
        df1_compare = self.df1.copy()
        df2_compare = self.df2.copy()
        
        # Apply comparison options
        if self.numbers_only.isChecked():
            # Try to convert columns to numeric for comparison
            try:
                df1_compare[col1] = pd.to_numeric(df1_compare[col1], errors='coerce')
                df2_compare[col2] = pd.to_numeric(df2_compare[col2], errors='coerce')
                
            except Exception as e:
                self.results_label.setText(f"Chyba převodu na čísla: {str(e)}")
                return
        elif not self.case_sensitive.isChecked():
            # Case insensitive comparison for strings
            try:
                df1_compare[col1] = df1_compare[col1].astype(str).str.lower()
                df2_compare[col2] = df2_compare[col2].astype(str).str.lower()
            except Exception as e:
                self.results_label.setText(f"Chyba v porovnání textu: {str(e)}")
                return
        
        # Find matches and differences
        # Convert numeric values with zero decimal places to int using lambda
        set1 = set(df1_compare[col1].dropna().apply(
            lambda x: int(x) if isinstance(x, (float, np.float64, int)) and x.is_integer() else x
        ))
        set2 = set(df2_compare[col2].dropna().apply(
            lambda x: int(x) if isinstance(x, (float, np.float64, int)) and x.is_integer() else x
        ))
        
        matches = set1.intersection(set2)
        only_in_df1 = set1.difference(set2)
        only_in_df2 = set2.difference(set1)
        
        # Display results
        self.results_label.setText(f"Shody: {len(matches)}, Pouze v první tabulce: {len(only_in_df1)}, Pouze v druhé tabulce: {len(only_in_df2)}, Celkem: {len(set1.union(set2))}")
        
        # Create a DataFrame with comparison results
        results_data = []
        
        # Add matches
        for value in matches:
            results_data.append({"Hodnota v prohledávaných sloupcích": value, "Výsledek porovnání": "Shoda", "Tabulka": "Obě"})
        
        # Add items only in file 1
        for value in only_in_df1:
            results_data.append({"Hodnota v prohledávaných sloupcích": value, "Výsledek porovnání": "Chybí", "Tabulka": "Pouze v tabulce 1"})
        
        # Add items only in file 2
        for value in only_in_df2:
            results_data.append({"Hodnota v prohledávaných sloupcích": value, "Výsledek porovnání": "Chybí", "Tabulka": "Pouze v tabulce 2"})
        
        # Display in table
        self.display_results(results_data)
    
    def display_results(self, results_data):
        self.results_table.clear()
        self.results_table.setRowCount(len(results_data))
        self.results_table.setColumnCount(3)
        self.results_table.setHorizontalHeaderLabels(["Hodnota v prohledávaných sloupcích", "Výsledek porovnání", "Tabulka"])
        
        for row, item in enumerate(results_data):
            self.results_table.setItem(row, 0, QTableWidgetItem(str(item["Hodnota v prohledávaných sloupcích"])))
            self.results_table.setItem(row, 1, QTableWidgetItem(item["Výsledek porovnání"]))
            self.results_table.setItem(row, 2, QTableWidgetItem(item["Tabulka"]))
            
            # Color code: green for matches, red for differences
            if item["Výsledek porovnání"] == "Shoda":
                self.results_table.item(row, 1).setBackground(Qt.green)
            else:
                self.results_table.item(row, 1).setBackground(Qt.red)
        
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.results_table.sortItems(1)  # Sort by status

    def export_results(self):
        if self.results_table.rowCount() == 0:
            self.results_label.setText("Žádné výsledky k exportu")
            return
        
        save_path, _ = QFileDialog.getSaveFileName(self, "Uložit výsledky", "", "Soubory Excelu (*.xlsx *.xls)")
        if save_path:
            results_df = pd.DataFrame(columns=["Hodnota v prohledávaných sloupcích", "Výsledek porovnání", "Tabulka"])
            for row in range(self.results_table.rowCount()):
                value = self.results_table.item(row, 0).text()
                result = self.results_table.item(row, 1).text()
                table = self.results_table.item(row, 2).text()
                new_row = pd.DataFrame({"Hodnota v prohledávaných sloupcích": [value], "Výsledek porovnání": [result], "Tabulka": [table]})
                results_df = pd.concat([results_df, new_row], ignore_index=True)
            
            results_df.to_excel(save_path, index=False)
            self.results_label.setText(f"Výsledky byly úspěšně uloženy do {os.path.basename(save_path)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelComparer()
    window.show()
    sys.exit(app.exec_())