import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QPushButton, QWidget, QFileDialog, QHBoxLayout, QSizePolicy, QComboBox, QButtonGroup
from PyQt5.QtCore import Qt
from PyQt5 import QtGui

import openpyxl
import sys
import pingouin as pg
import pandas as pd
from company_ontology_model import *



sys.path.append("d:/UIT-peR/CSMHTT_QH/Exercises/RelKnow_final/source_code")
           
class ExcelApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.number_quarters = 4

        self.setWindowTitle("Hệ thống dự báo tài chính doanh nghiệp - CS2231")

        # Create a QWidget to hold everything
        widget = QWidget()
        self.setCentralWidget(widget)

        # Create a QVBoxLayout to layout our widgets
        layout = QVBoxLayout()
        widget.setLayout(layout)

        # Create a QHBoxLayout for the buttons
        button_layout = QHBoxLayout()
        button_layout.setAlignment(Qt.AlignLeft)
        button_layout.setContentsMargins(0, 0, 10, 0)

        # Create buttons
        self.import_button = QPushButton("Nhập file")
        self.import_button.clicked.connect(self.on_import)
        self.import_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        button_layout.addWidget(self.import_button)

        self.calculate_button = QPushButton("Tính toán")
        self.calculate_button.clicked.connect(self.on_calculate)
        self.calculate_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        button_layout.addWidget(self.calculate_button)

        self.export_button = QPushButton("Xuất file")
        self.export_button.clicked.connect(self.on_export)
        self.export_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        button_layout.addWidget(self.export_button)

        self.quit_button = QPushButton("Quit")
        self.quit_button.clicked.connect(self.close)
        self.quit_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        button_layout.addWidget(self.quit_button)

        # Add the button layout to the main layout
        layout.addLayout(button_layout)

        # Create QTableWidget (grid)
        self.table = QTableWidget(0, 0)  # 0 rows, 0 columns initially
        self.table.horizontalHeader().setStyleSheet("color: blue; background-color: white");
        # Create a QHBoxLayout for the combo box and the "Remove Company" button
        combo_layout = QHBoxLayout()
        combo_layout.setAlignment(Qt.AlignLeft)  # Align to the left
        combo_layout.setContentsMargins(0, 0, 10, 0)  # Set right margin to 10

        # Create a QComboBox for company selection
        self.company_combo = QComboBox()
        self.company_combo.currentTextChanged.connect(self.on_company_selected)
        self.company_combo.setFixedWidth(500)  # Set the width of the combo box to 300
        combo_layout.addWidget(self.company_combo)

        # Create a "Remove Company" button
        self.remove_company_button = QPushButton("Xóa")
        self.remove_company_button.clicked.connect(self.on_remove_clicked)
        self.remove_company_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        combo_layout.addWidget(self.remove_company_button)

        # Add the QHBoxLayout to the main layout
        layout.addLayout(combo_layout)

        # Populate the company combo box
        self.refresh_company_combo()

        layout.addWidget(self.table)

        # Set the window width and height
        self.resize(1500, 800)  # Replace 800 and 600 with your desired width and height

    def on_company_selected(self, company_name):
        self.selected_company = f"INPUT_{company_name}.xlsx"
        print(f"Selected company: {self.selected_company}")

        # Clear the table
        if hasattr(self, 'table') and self.table is not None:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)

        # Load the data file
        file_path = os.path.join("data", self.selected_company)
        if os.path.exists(file_path):
            self.load_excel_data(file_path)
        else:
            print(f"File not found: {file_path}")

    def on_remove_clicked(self):
        # Get the currently selected company
        company_name = self.company_combo.currentText()

        # Find the index of the selected company in the combo box
        index = self.company_combo.findText(company_name)

        # If the company is found, remove it
        if index >= 0:
            self.company_combo.removeItem(index)

            # Remove the corresponding file
            file_path = os.path.join("data", f"INPUT_{company_name}.xlsx")
            if os.path.exists(file_path):
                os.remove(file_path)
    
    def on_import(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "", "Excel files (*.xlsx)")
        if file_path:
            # Get the filename from the selected path
            filename = os.path.basename(file_path)

            # Clear the table
            self.table.setRowCount(0)
            self.table.setColumnCount(0)

            # Save the file to the "data" folder with the same name
            data_folder = "data"
            os.makedirs(data_folder, exist_ok=True)  # Create the "data" folder if it doesn't exist
            save_path = os.path.join(data_folder, filename)
            self.save_excel_data(file_path, save_path)  # Call the save method with the source and destination paths

            # Load the data from the saved file
            self.load_excel_data(save_path)

            # Refresh the company combo box
            self.refresh_company_combo()

            # Update the selected company based on the filename
            if filename.startswith("INPUT_") and filename.endswith(".xlsx"):
                company_name = filename[6:-5]
                self.company_combo.setCurrentText(company_name)
                self.selected_company = filename
            else:
                print("Invalid filename format. Expected: INPUT_XXX.xlsx")

    def refresh_company_combo(self):
        self.company_combo.clear()
        data_folder = "data"
        print(os.path.abspath('.').split(os.path.sep)[0]+os.path.sep)
        for file_name in os.listdir(data_folder):
            if file_name.startswith("INPUT_") and file_name.endswith(".xlsx"):
                company_name = file_name[6:-5]
                self.company_combo.addItem(company_name)

    def save_excel_data(self, source_path, dest_path):
        if self.selected_company:
            import shutil
            shutil.copy(source_path, dest_path)
        else:
            print("Please select a company first.")

    def load_excel_data(self, file_path):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        data = list(sheet.values)

        if data:
            # Get column headers from the first row
            headers = data[0]
            self.number_quarters = len(headers)-4

            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)
            
            # Set the width of the second column to 400px
            self.table.setColumnWidth(1, 600)

            # Insert data rows
            for (ix,row) in enumerate(data[1:]):
                row_num = self.table.rowCount()
                self.table.insertRow(row_num)
                for col_num, cell in enumerate(row):
                    if cell is None:
                        cell = ""
                    if col_num>1 and cell !='' and isinstance(cell, float):
                        cell=str(round(float(cell),2))
                    item = QTableWidgetItem(str(cell))
                    if (ix ==0):   
                        item.setBackground(QtGui.QColor(173,216,230))
                    self.table.setItem(row_num, col_num, item)
                    print(f"Row: {row_num}, Column: {col_num}, Item: {item.text()}")


    def on_calculate_additional(self):
        A1_Q1 = "{:.0f}%".format(float(self.table.item(2, 2).text()) / float(self.table.item(8, 2).text()) * 100)
        A1_Q2 = "{:.0f}%".format(float(self.table.item(2, 3).text()) / float(self.table.item(8, 3).text()) * 100)
        A1_Q3 = "{:.0f}%".format(float(self.table.item(2, 4).text()) / float(self.table.item(8, 4).text()) * 100)
        A1_Q4 = "{:.0f}%".format(float(self.table.item(2, 5).text()) / float(self.table.item(8, 5).text()) * 100)
        A1_AVG = "{:.0f}%".format((float(A1_Q1[:-1]) + float(A1_Q2[:-1]) + float(A1_Q3[:-1]) + float(A1_Q4[:-1])) / 4)

        A2_Q1 = "{:.0f}%".format((float(self.table.item(2, 2).text()) - float(self.table.item(4, 2).text())) / float(self.table.item(8, 2).text()) * 100)
        A2_Q2 = "{:.0f}%".format((float(self.table.item(2, 3).text()) - float(self.table.item(4, 3).text())) / float(self.table.item(8, 3).text()) * 100)
        A2_Q3 = "{:.0f}%".format((float(self.table.item(2, 4).text()) - float(self.table.item(4, 4).text())) / float(self.table.item(8, 4).text()) * 100)
        A2_Q4 = "{:.0f}%".format((float(self.table.item(2, 5).text()) - float(self.table.item(4, 5).text())) / float(self.table.item(8, 5).text()) * 100)
        A2_AVG = "{:.0f}%".format((float(A2_Q1[:-1]) + float(A2_Q2[:-1]) + float(A2_Q3[:-1]) + float(A2_Q4[:-1])) / 4)

        A3_Q1 = "{:.0f}%".format(float(self.table.item(3, 2).text()) / float(self.table.item(8, 2).text()) * 100)
        A3_Q2 = "{:.0f}%".format(float(self.table.item(3, 3).text()) / float(self.table.item(8, 3).text()) * 100)
        A3_Q3 = "{:.0f}%".format(float(self.table.item(3, 4).text()) / float(self.table.item(8, 4).text()) * 100)
        A3_Q4 = "{:.0f}%".format(float(self.table.item(3, 5).text()) / float(self.table.item(8, 5).text()) * 100)
        A3_AVG = "{:.0f}%".format((float(A3_Q1[:-1]) + float(A3_Q2[:-1]) + float(A3_Q3[:-1]) + float(A3_Q4[:-1])) / 4)

        B1_Q1 = "{:.0f}%".format(float(self.table.item(7, 2).text()) / float(self.table.item(1, 2).text()) * 100)
        B1_Q2 = "{:.0f}%".format(float(self.table.item(7, 3).text()) / float(self.table.item(1, 3).text()) * 100)
        B1_Q3 = "{:.0f}%".format(float(self.table.item(7, 4).text()) / float(self.table.item(1, 4).text()) * 100)
        B1_Q4 = "{:.0f}%".format(float(self.table.item(7, 5).text()) / float(self.table.item(1, 5).text()) * 100)
        B1_AVG = "{:.0f}%".format((float(B1_Q1[:-1]) + float(B1_Q2[:-1]) + float(B1_Q3[:-1]) + float(B1_Q4[:-1])) / 4)

        B2_Q1 = "{:.0f}%".format(float(self.table.item(7, 2).text()) / float(self.table.item(11, 2).text()) * 100)
        B2_Q2 = "{:.0f}%".format(float(self.table.item(7, 3).text()) / float(self.table.item(11, 3).text()) * 100)
        B2_Q3 = "{:.0f}%".format(float(self.table.item(7, 4).text()) / float(self.table.item(11, 4).text()) * 100)
        B2_Q4 = "{:.0f}%".format(float(self.table.item(7, 5).text()) / float(self.table.item(11, 5).text()) * 100)
        B2_AVG = "{:.0f}%".format((float(B2_Q1[:-1]) + float(B2_Q2[:-1]) + float(B2_Q3[:-1]) + float(B2_Q4[:-1])) / 4)

        B3_Q1 = "{:.0f}%".format(float(self.table.item(16, 2).text()) / float(self.table.item(15, 2).text()) * 100)
        B3_Q2 = "{:.0f}%".format(float(self.table.item(16, 3).text()) / float(self.table.item(15, 3).text()) * 100)
        B3_Q3 = "{:.0f}%".format(float(self.table.item(16, 4).text()) / float(self.table.item(15, 4).text()) * 100)
        B3_Q4 = "{:.0f}%".format(float(self.table.item(16, 5).text()) / float(self.table.item(15, 5).text()) * 100)
        B3_AVG = "{:.0f}%".format((float(B3_Q1[:-1]) + float(B3_Q2[:-1]) + float(B3_Q3[:-1]) + float(B3_Q4[:-1])) / 4)

        C1_Q1 = "{:.0f}%".format(float(self.table.item(13, 2).text()) / float(self.table.item(4, 6).text()) * 100)
        C1_Q2 = "{:.0f}%".format(float(self.table.item(13, 3).text()) / float(self.table.item(4, 6).text()) * 100)
        C1_Q3 = "{:.0f}%".format(float(self.table.item(13, 4).text()) / float(self.table.item(4, 6).text()) * 100)
        C1_Q4 = "{:.0f}%".format(float(self.table.item(13, 5).text()) / float(self.table.item(4, 6).text()) * 100)
        C1_AVG = "{:.0f}%".format((float(C1_Q1[:-1]) + float(C1_Q2[:-1]) + float(C1_Q3[:-1]) + float(C1_Q4[:-1])) / 4)

        C2_Q1 = "{:.0f}%".format(float(self.table.item(9, 2).text()) / float(self.table.item(12, 6).text()) * 100)
        C2_Q2 = "{:.0f}%".format(float(self.table.item(9, 3).text()) / float(self.table.item(12, 6).text()) * 100)
        C2_Q3 = "{:.0f}%".format(float(self.table.item(9, 4).text()) / float(self.table.item(12, 6).text()) * 100)
        C2_Q4 = "{:.0f}%".format(float(self.table.item(9, 5).text()) / float(self.table.item(12, 6).text()) * 100)
        C2_AVG = "{:.0f}%".format((float(C2_Q1[:-1]) + float(C2_Q2[:-1]) + float(C2_Q3[:-1]) + float(C2_Q4[:-1])) / 4)

        C3_Q1 = "{:.0f}%".format(float(self.table.item(12, 2).text()) / float(self.table.item(5, 6).text()) * 100)
        C3_Q2 = "{:.0f}%".format(float(self.table.item(12, 3).text()) / float(self.table.item(5, 6).text()) * 100)
        C3_Q3 = "{:.0f}%".format(float(self.table.item(12, 4).text()) / float(self.table.item(5, 6).text()) * 100)
        C3_Q4 = "{:.0f}%".format(float(self.table.item(12, 5).text()) / float(self.table.item(5, 6).text()) * 100)
        C3_AVG = "{:.0f}%".format((float(C3_Q1[:-1]) + float(C3_Q2[:-1]) + float(C3_Q3[:-1]) + float(C3_Q4[:-1])) / 4)

        D1_Q1 = "{:.0f}%".format(float(self.table.item(17, 2).text()) / float(self.table.item(12, 2).text()) * 100)
        D1_Q2 = "{:.0f}%".format(float(self.table.item(17, 3).text()) / float(self.table.item(12, 3).text()) * 100)
        D1_Q3 = "{:.0f}%".format(float(self.table.item(17, 4).text()) / float(self.table.item(12, 4).text()) * 100)
        D1_Q4 = "{:.0f}%".format(float(self.table.item(17, 5).text()) / float(self.table.item(12, 5).text()) * 100)
        D1_AVG = "{:.0f}%".format((float(D1_Q1[:-1]) + float(D1_Q2[:-1]) + float(D1_Q3[:-1]) + float(D1_Q4[:-1])) / 4)

        D2_Q1 = "{:.0f}%".format(float(self.table.item(16, 2).text()) / float(self.table.item(1, 6).text()) * 100)
        D2_Q2 = "{:.0f}%".format(float(self.table.item(16, 3).text()) / float(self.table.item(1, 6).text()) * 100)
        D2_Q3 = "{:.0f}%".format(float(self.table.item(16, 4).text()) / float(self.table.item(1, 6).text()) * 100)
        D2_Q4 = "{:.0f}%".format(float(self.table.item(16, 5).text()) / float(self.table.item(1, 6).text()) * 100)
        D2_AVG = "{:.0f}%".format((float(D2_Q1[:-1]) + float(D2_Q2[:-1]) + float(D2_Q3[:-1]) + float(D2_Q4[:-1])) / 4)

        D3_Q1 = "{:.0f}%".format(float(self.table.item(17, 2).text()) / float(self.table.item(1, 6).text()) * 100)
        D3_Q2 = "{:.0f}%".format(float(self.table.item(17, 3).text()) / float(self.table.item(1, 6).text()) * 100)
        D3_Q3 = "{:.0f}%".format(float(self.table.item(17, 4).text()) / float(self.table.item(1, 6).text()) * 100)
        D3_Q4 = "{:.0f}%".format(float(self.table.item(17, 5).text()) / float(self.table.item(1, 6).text()) * 100)
        D3_AVG = "{:.0f}%".format((float(D3_Q1[:-1]) + float(D3_Q2[:-1]) + float(D3_Q3[:-1]) + float(D3_Q4[:-1])) / 4)

        # Add more rows like in the provided data
        additional_rows = [
            ["",  "CÁC HỆ SỐ PHÂN TÍCH CHÍNH", "", "", "", "","",""],
            ["",  "", "", "", "", "","",""],
            ["19", "1. Khả năng thanh toán", "", "", "", ""],
            ["20", "A1: Tỷ số khả năng thanh toán hiện thời = Tài sản ngắn hạn/ Nợ ngắn hạn", A1_Q1, A1_Q2, A1_Q3, A1_Q4, A1_AVG],
            ["21", "A2: Tỷ số khả năng thanh toán nhanh = (Tài sản ngắn hạn – Tồn kho)/ Nợ ngắn hạn", A2_Q1, A2_Q2, A2_Q3, A2_Q4, A2_AVG],
            ["22", "A3: Tỷ số khả năng thanh toán tức thời = Tiền/ Nợ ngắn hạn", A3_Q1, A3_Q2, A3_Q3, A3_Q4, A3_AVG],
            ["23", "B: Khả năng cân đối vốn", "", "", "", ""],
            ["24", "B1: Tỷ số nợ trên tổng tài sản (hệ số nợ) = Nợ phải trả/ Tổng tài sản", B1_Q1, B1_Q2, B1_Q3, B1_Q4, B1_AVG],
            ["25", "B2: Tỷ số Nợ phải trả trên Vốn chủ sở hữu = Nợ phải trả/ Vốn chủ sở hữu", B2_Q1, B2_Q2, B2_Q3, B2_Q4, B2_AVG],
            ["26", "B3: Tỷ số khả năng thanh toán lãi vay (TIE) = EBIT (lợi nhuận trước lãi vay và thuế)/ Lãi vay.", B3_Q1, B3_Q2, B3_Q3, B3_Q4, B3_AVG],
            ["27", "C: Hiệu quả hoạt động", "", "", "", ""],
            ["28", "C1: Vòng quay hàng tồn kho = Giá vốn hàng bán/ Hàng tồn kho bình quân", C1_Q1, C1_Q2, C1_Q3, C1_Q4, C1_AVG],
            ["29", "C2: Kỳ thu tiền trung bình = Khoản phải thu ngắn hạn bình quân/ Doanh thu thuần bình quân", C2_Q1, C2_Q2, C2_Q3, C2_Q4, C2_AVG],
            ["30", "C3: Vòng quay tài sản cố định = Doanh thu thuần/ Tài sản cố định ròng bình quân", C3_Q1, C3_Q2, C3_Q3, C3_Q4, C3_AVG],
            ["31", "D. Khả năng sinh lợi", "", "", "", ""],
            ["32", "D1: Tỷ suất doanh lợi doanh thu (ROS) = Lợi nhuận sau thuế/ Doanh thu thuần", D1_Q1, D1_Q2, D1_Q3, D1_Q4, D1_AVG],
            ["33", "D2: Tỷ số khả năng sinh lời cơ bản của tài sản = EBIT (lợi nhuận trước thuế và lãi vay)/ Tổng tài sản bình quân", D2_Q1, D2_Q2, D2_Q3, D2_Q4, D2_AVG],
            ["34", "D3: Tỷ suất doanh lợi tổng tài sản (ROA) = Lợi nhuận sau thuế/ Tổng tài sản bình quân", D3_Q1, D3_Q2, D3_Q3, D3_Q4, D3_AVG]
        ]

        for (ix,row) in enumerate(additional_rows):
            row_num = self.table.rowCount()
            self.table.insertRow(row_num)
            for col_num, cell in enumerate(row):
                item = QTableWidgetItem(str(cell))
                if (ix ==0):   
                    item.setBackground(QtGui.QColor(173,216,230))
                self.table.setItem(row_num, col_num, item)

    def row_quater_list(self,index)->list:
        arr=[]
        for i in range(self.number_quarters-2):
            arr.append(float(self.table.item(index, i+2).text().strip('%'))) 
        return arr

    def on_calculate_cronch(self):
        df = pd.DataFrame(
            {
                'A1': [float(self.table.item(21, 2).text().strip('%')), float(self.table.item(21, 3).text().strip('%')), float(self.table.item(21, 4).text().strip('%')), float(self.table.item(21, 5).text().strip('%'))],
                'A2': [float(self.table.item(22, 2).text().strip('%')), float(self.table.item(22, 3).text().strip('%')), float(self.table.item(22, 4).text().strip('%')), float(self.table.item(22, 5).text().strip('%'))],
                'A3': [float(self.table.item(23, 2).text().strip('%')), float(self.table.item(23, 3).text().strip('%')), float(self.table.item(23, 4).text().strip('%')), float(self.table.item(23, 5).text().strip('%'))],
                'B1': [float(self.table.item(25, 2).text().strip('%')), float(self.table.item(25, 3).text().strip('%')), float(self.table.item(25, 4).text().strip('%')), float(self.table.item(25, 5).text().strip('%'))],
                'B2': [float(self.table.item(26, 2).text().strip('%')), float(self.table.item(26, 3).text().strip('%')), float(self.table.item(26, 4).text().strip('%')), float(self.table.item(26, 5).text().strip('%'))],
                'B3': [float(self.table.item(27, 2).text().strip('%')), float(self.table.item(27, 3).text().strip('%')), float(self.table.item(27, 4).text().strip('%')), float(self.table.item(27, 5).text().strip('%'))],
                'C1': [float(self.table.item(29, 2).text().strip('%')), float(self.table.item(29, 3).text().strip('%')), float(self.table.item(29, 4).text().strip('%')), float(self.table.item(29, 5).text().strip('%'))],
                'C2': [float(self.table.item(30, 2).text().strip('%')), float(self.table.item(30, 3).text().strip('%')), float(self.table.item(30, 4).text().strip('%')), float(self.table.item(30, 5).text().strip('%'))],
                'C3': [float(self.table.item(31, 2).text().strip('%')), float(self.table.item(31, 3).text().strip('%')), float(self.table.item(31, 4).text().strip('%')), float(self.table.item(31, 5).text().strip('%'))],
                'D1': [float(self.table.item(33, 2).text().strip('%')), float(self.table.item(33, 3).text().strip('%')), float(self.table.item(33, 4).text().strip('%')), float(self.table.item(33, 5).text().strip('%'))],
                'D2': [float(self.table.item(34, 2).text().strip('%')), float(self.table.item(34, 3).text().strip('%')), float(self.table.item(34, 4).text().strip('%')), float(self.table.item(34, 5).text().strip('%'))],
                'D3': [float(self.table.item(35, 2).text().strip('%')), float(self.table.item(35, 3).text().strip('%')), float(self.table.item(35, 4).text().strip('%')), float(self.table.item(35, 5).text().strip('%'))]
            }
        )
        # data_raw={
        #         'A1': self.row_quater_list(21) ,
        #          'A2': self.row_quater_list(22) ,
        #           'A3': self.row_quater_list(23) ,
        #            'B1': self.row_quater_list(25) ,
        #             'B2': self.row_quater_list(26) ,
        #              'B3': self.row_quater_list(27) ,
        #               'C1': self.row_quater_list(29) ,
        #                'C2': self.row_quater_list(30) ,
        #                 'C3': self.row_quater_list(31) ,
        #                  'D1': self.row_quater_list(33) ,
        #                   'D2': self.row_quater_list(34) ,
        #                    'D3': self.row_quater_list(35) 
        #     }

        # df = pd.DataFrame(data_raw)
        # Calculate Cronbach's alpha and round to 2 decimal places
        alpha_A = round(pg.cronbach_alpha(df[['A1', 'A2', 'A3']])[0], 2)
        alpha_B = round(pg.cronbach_alpha(df[['B1', 'B2', 'B3']])[0], 2)
        alpha_C = round(pg.cronbach_alpha(df[['C1', 'C2', 'C3']])[0], 2)
        alpha_D = round(pg.cronbach_alpha(df[['D1', 'D2', 'D3']])[0], 2)
        
        # Set the alpha values in the table
        self.table.setItem(20, 7, QTableWidgetItem(str(alpha_A)))
        self.table.setItem(24, 7, QTableWidgetItem(str(alpha_B)))
        self.table.setItem(28, 7, QTableWidgetItem(str(alpha_C)))
        self.table.setItem(32, 7, QTableWidgetItem(str(alpha_D)))

    def get_comment_with_value(self,value):
        if value>=0.9:
             return [" thang đo xuất sắc",True]
        if value>=0.8:
            return [" thang đo tốt",True]
        if value>=0.7:
            return [" thang đo sử dụng được",True]
        if value>=0.6:
            return [" thang đo sử dụng được trong bối cảnh nghiên cứu mới",True]
        if value>=0.5:
            return [" thang đo kém, cần xem xét lại",True]          
        return [" thang đo không được chấp nhận",False]               
    def on_calculate_cronch_comments(self):
        cronbach_alphaA="Khả năng thanh toán: "+ self.table.item(20, 7).text().strip('%')+", "+self.get_comment_with_value(float(self.table.item(20, 7).text().strip('%')))[0]
        cronbach_alphaB="Khả năng cân đối vốn: "+ self.table.item(24, 7).text().strip('%')+", "+self.get_comment_with_value(float(self.table.item(24, 7).text().strip('%')))[0]
        cronbach_alphaC="Hiệu quả hoạt động: "+ self.table.item(28, 7).text().strip('%')+", "+self.get_comment_with_value(float(self.table.item(28, 7).text().strip('%')))[0]
        cronbach_alphaD="Khả năng sinh lợi: "+ self.table.item(32, 7).text().strip('%')+", "+self.get_comment_with_value(float(self.table.item(32, 7).text().strip('%')))[0]
        list_color=[]
        for id in range(4):
            list_color.append(self.get_comment_with_value(float(self.table.item(20+4*id, 7).text().strip('%')))[1])
         # Add more rows like in the provided data
        additional_rows = [
            ["", "", "", "", "", "", "", ""],
            ["", "KIỂM ĐỊNH ĐỘ TIN CẬY THANG ĐO CRONBACH'S ALPHA :", "", "", "", "", "", ""],
            ["35", cronbach_alphaA, "", "", "", "", "", ""],
            ["36", cronbach_alphaB, "", "", "", "", "", ""],
            ["37", cronbach_alphaC, "", "", "", "", "", ""],
            ["38", cronbach_alphaD, "", "", "", "", "", ""],

        ]

        for (ix,row) in enumerate(additional_rows):
            row_num = self.table.rowCount()
            self.table.insertRow(row_num)
            for col_num, cell in enumerate(row):
                item = QTableWidgetItem(str(cell))
                if (ix>1) and (col_num==1) and (list_color[ix-2]==False):             
                    item.setBackground(QtGui.QColor(255, 128, 128))
                if (ix ==1):   
                    item.setBackground(QtGui.QColor(173,216,230))
                self.table.setItem(row_num, col_num, item)
    
    
    def on_calculate_final_comments(self):
        input_dic={"total_assets": float(self.table.item(1, 5).text().strip('%')), \
                    "short_term_assets":float(self.table.item(2, 5).text().strip('%')), \
                      "cash":float(self.table.item(3, 5).text().strip('%')), \
                       "inventory":float(self.table.item(4, 5).text().strip('%')), \
                       "long_term_assets":float(self.table.item(5, 5).text().strip('%')), \
                       "total_capital":float(self.table.item(6, 5).text().strip('%')), \
                        "liabilities":float(self.table.item(7, 5).text().strip('%')), \
                         "short_term_debt":float(self.table.item(8, 5).text().strip('%')), \
                         "short_term_receivables":float(self.table.item(9, 5).text().strip('%')), \
                        "long_term_debt":float(self.table.item(10, 5).text().strip('%')), \
                        "equity":float(self.table.item(11, 5).text().strip('%')), \
                        "net_revenue":float(self.table.item(12, 5).text().strip('%')), \
                          "cost_of_sales":float(self.table.item(13, 5).text().strip('%')), \
                          "profit_before_tax":float(self.table.item(14, 5).text().strip('%')), \
                         "interest_expenses":float(self.table.item(15, 5).text().strip('%')), \
                        "ebit":float(self.table.item(16, 5).text().strip('%')), \
                         "profit_after_tax":float(self.table.item(17, 5).text().strip('%')),\
                         
                         "average_inventory":float(self.table.item(4, 6).text().strip('%')),\
                         "average_short_term_receivables":float(self.table.item(9, 6).text().strip('%')),\
                         "average_net_revenue":float(self.table.item(12, 6).text().strip('%')),\
                         "average_net_fixed_assets":float(self.table.item(5, 6).text().strip('%')),\
                         "average_total_assets":float(self.table.item(1, 6).text().strip('%'))
                      }
        
        print(float(self.table.item(1, 5).text().strip('%')))
        comments=print_test(input_dic)
        list_color=[]
        for id in range(4):
            list_color.append(self.get_comment_with_value(float(self.table.item(20+4*id, 7).text().strip('%')))[1])

         # Add more rows like in the provided data
        additional_rows = [
        ]
        for row in comments:
            additional_rows.append( ["", row, "", "", "", "", "", ""])
        additional_rows.append(["", "", "", "", "", "", "", ""])



        for (ix,row) in enumerate(additional_rows):
            row_num = self.table.rowCount()
            self.table.insertRow(row_num)
            for col_num, cell in enumerate(row):
                item = QTableWidgetItem(str(cell))
                if (ix>1) and (col_num==1) and (int((ix-2)/5)<4) and (list_color[int((ix-2)/5)]==False):             
                    item.setBackground(QtGui.QColor(255, 128, 128))
                if (ix ==1):   
                    item.setBackground(QtGui.QColor(173,216,230))
                self.table.setItem(row_num, col_num, item)

    def refresh_the_table(self):
        # Refresh the table with loaded data and remove the calculated rows
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.load_excel_data(os.path.join("data", self.selected_company))

    def on_calculate(self):
        self.refresh_the_table()
        self.on_calculate_additional()
        self.on_calculate_cronch()
        self.on_calculate_cronch_comments()
        self.on_calculate_final_comments()
    
    def on_export(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel file", "", "Excel files (*.xlsx)")
        if file_path:
            workbook = openpyxl.Workbook()
            sheet = workbook.active

        # Write column headers
            headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
            sheet.append(headers)

        # Write data rows
            for row_num in range(self.table.rowCount()):
                row_data = [self.table.item(row_num, col_num).text() if self.table.item(row_num, col_num) is not None else '' for col_num in range(self.table.columnCount())]
                sheet.append(row_data)
            workbook.save(file_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = ExcelApp()
    window.show()

    sys.exit(app.exec_())