import sys
import pandas as pd
import os
import shutil
import json
import re
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QPlainTextEdit, QMessageBox, QHeaderView, QFrame
)
from PyQt5.QtCore import Qt


# ==========================================================
# 🔍 FUNGSI DEBUG API
# ==========================================================
def get_api_data(url):
    try:
        print("\n=== MENGAMBIL DATA API ===")
        print(f"URL: {url}")
        response = requests.get(url, verify=False, timeout=15)

        print("\n=== STATUS CODE ===")
        print(response.status_code)

        print("\n=== RAW RESPONSE (TEXT) ===")
        print(response.text)

        print("\n=== MENCOBA PARSE JSON ===")
        data = json.loads(response.text)

        print("=== JSON PARSED BERHASIL ===")
        return data

    except Exception as e:
        print("\n❌ ERROR Saat Ambil API:", e)
        return None


# ==========================================================
# 🚀 APLIKASI UTAMA
# ==========================================================
class ExcelToTxtApp(QWidget):
    def __init__(self):
        super().__init__()

        self.api_connected = False
        self.berat_haspel = {}
        self.df = None

        self.init_ui()

        # 🔥 FULLSCREEN SAAT APLIKASI DIMULAI
        self.showFullScreen()

        # Load API
        self.load_haspel_from_api()

    # ==========================================================
    # 🔥 ESC → KELUAR FULLSCREEN
    # ==========================================================
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.showNormal()

    # ==========================================================
    # 🔥 LOAD API & SIMPAN KE DICTIONARY
    # ==========================================================
    def load_haspel_from_api(self):
        url = "https://pl.jembo.com/index.php/PackingListStandar/get_berat_haspel_all"
        data = get_api_data(url)

        if data is None:
            self.api_connected = False
            self.berat_haspel = {}
            self.append_log("❌ API tidak terhubung, Silahkan connect ke WiFi JEMBO")
            return

        try:
            self.berat_haspel = {
                kode: {"berat": float(info.get("berat", 0))}
                for kode, info in data.items()
            }
            self.api_connected = True
            self.append_log("✅ API Connected (data haspel berhasil dimuat).")

        except Exception as e:
            self.api_connected = False
            self.berat_haspel = {}
            self.append_log(f"❌ Gagal parsing data API: {e}")

    # ==========================================================
    # UI
    # ==========================================================
    def init_ui(self):
        self.setWindowTitle("MINI TOOLS")
        self.setGeometry(100, 100, 1200, 700)

        main_layout = QVBoxLayout()

        # ========== CARD 1 ==========
        card_1 = QFrame()
        card_layout_1 = QVBoxLayout()

        excel_layout = QHBoxLayout()
        self.label_excel = QLabel("Select Excel File:")
        self.entry_excel = QLineEdit()
        self.entry_excel.setReadOnly(True)
        self.button_browse_excel = QPushButton("Browse")
        self.button_browse_excel.clicked.connect(self.browse_excel)
        excel_layout.addWidget(self.label_excel)
        excel_layout.addWidget(self.entry_excel)
        excel_layout.addWidget(self.button_browse_excel)

        output_layout = QHBoxLayout()
        self.label_output = QLabel("Select Output Folder:")
        self.entry_output = QLineEdit()
        self.entry_output.setReadOnly(True)
        self.button_browse_output = QPushButton("Browse")
        self.button_browse_output.clicked.connect(self.browse_output)
        output_layout.addWidget(self.label_output)
        output_layout.addWidget(self.entry_output)
        output_layout.addWidget(self.button_browse_output)

        self.button_load = QPushButton("Load Data")
        self.button_load.clicked.connect(self.load_data)

        card_layout_1.addLayout(excel_layout)
        card_layout_1.addLayout(output_layout)
        card_layout_1.addWidget(self.button_load)
        card_1.setLayout(card_layout_1)

        # ========== CARD 2 (LOT) ==========
        card_2 = QFrame()
        card_layout_2 = QVBoxLayout()

        select_layout = QHBoxLayout()
        self.button_select_all = QPushButton("Select All")
        self.button_select_all.clicked.connect(self.select_all)
        self.button_deselect_all = QPushButton("Deselect All")
        self.button_deselect_all.clicked.connect(self.deselect_all)
        select_layout.addWidget(self.button_select_all)
        select_layout.addWidget(self.button_deselect_all)

        search_layout = QHBoxLayout()
        self.label_search = QLabel("Search Lot:")
        self.search_box = QLineEdit()
        self.search_box.textChanged.connect(self.search_table)
        search_layout.addWidget(self.label_search)
        search_layout.addWidget(self.search_box)

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(1)
        self.table_widget.setHorizontalHeaderLabels(["Lot"])
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_widget.setSelectionMode(QAbstractItemView.MultiSelection)
        self.table_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

        self.button_process = QPushButton("Process")
        self.button_process.clicked.connect(self.process_excel)

        card_layout_2.addLayout(select_layout)
        card_layout_2.addLayout(search_layout)
        card_layout_2.addWidget(self.table_widget)
        card_layout_2.addWidget(self.button_process)
        card_2.setLayout(card_layout_2)

        # ========== CARD 3 (LOG) ==========
        card_3 = QFrame()
        card_layout_3 = QVBoxLayout()

        self.log_label = QLabel("Log:")
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.clear_log_button = QPushButton("Clear Log")
        self.clear_log_button.clicked.connect(self.clear_log)

        card_layout_3.addWidget(self.log_label)
        card_layout_3.addWidget(self.log_text)
        card_layout_3.addWidget(self.clear_log_button)
        card_3.setLayout(card_layout_3)

        # MAIN LAYOUT
        main_layout.addWidget(card_1)

        lot_log_layout = QHBoxLayout()
        lot_log_layout.addWidget(card_2, 70)
        lot_log_layout.addWidget(card_3, 30)
        main_layout.addLayout(lot_log_layout)

        self.setLayout(main_layout)

    # UTIL MENUS
    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.entry_excel.setText(file_path)

    def browse_output(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder_path:
            self.entry_output.setText(folder_path)

    def append_log(self, msg):
        self.log_text.appendPlainText(msg)

    def clear_log(self):
        self.log_text.clear()

    def sanitize_filename(self, name):
        return "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in name).strip()

    # LOAD DATA EXCEL
    def load_data(self):
        try:
            excel_file = self.entry_excel.text()
            self.df = pd.read_excel(excel_file)
            self.df.columns = self.df.columns.str.strip()

            lots = self.df["Lot"].unique()
            self.table_widget.setRowCount(0)

            for lot in lots:
                row = self.table_widget.rowCount()
                self.table_widget.insertRow(row)
                self.table_widget.setItem(row, 0, QTableWidgetItem(str(lot)))

            self.append_log("Excel loaded successfully.")

        except Exception as e:
            self.append_log(f"Error: {e}")

    # PROCESS
    def process_excel(self):
        if not self.api_connected:
            self.append_log("❌ API tidak terhubung → proses dibatalkan.")
            return

        try:
            output_folder = self.entry_output.text()
            selected_rows = self.table_widget.selectedIndexes()
            selected_lots = list({self.table_widget.item(i.row(), 0).text() for i in selected_rows})

            if not selected_lots:
                self.append_log("No lot selected.")
                return

            for lot in selected_lots:
                group = self.df[self.df["Lot"] == lot]

                lot_folder = os.path.join(output_folder, self.sanitize_filename(str(lot)))
                if os.path.exists(lot_folder):
                    shutil.rmtree(lot_folder)
                os.makedirs(lot_folder, exist_ok=True)

                for _, row in group.iterrows():
                    description = row.get("Description", "")
                    quantity_lot = row.get("Quantity Lot", 0)
                    weight = row.get("Weight", 0)
                    customer_name = row.get("CustomerName", "")
                    standard_desc = row.get("Standard Description", "")
                    spec = row.get("Spec", "")

                    match = re.search(r"\d+", str(standard_desc))
                    code_num = match.group() if match else "000"

                    berat_haspel = float(self.berat_haspel.get(code_num, {}).get("berat", 0))

                    netto = int(round(quantity_lot * weight))
                    gross = int(round(netto + berat_haspel))

                    files = [
                        (os.path.join(lot_folder, f"01-{lot}-No.Drum.txt"), f"{code_num}-{lot}\n"),
                        (os.path.join(lot_folder, f"02-{lot}-Customer.txt"), f"{customer_name}\n"),
                        (os.path.join(lot_folder, f"03-{lot}-Description.txt"), f"{description}\n"),
                        (os.path.join(lot_folder, f"04-{lot}-Length.txt"), f"LENGTH : {quantity_lot} M\n"),
                        (os.path.join(lot_folder, f"05-{lot}-Netto.txt"), f"NETTO : {netto} KG\n"),
                        (os.path.join(lot_folder, f"06-{lot}-Gross.txt"), f"GROSS : {gross} KG\n"),
                        (os.path.join(lot_folder, f"07-{lot}-Roll_This_Way.txt"), "Roll This Way →\n"),
                        (os.path.join(lot_folder, f"08-{lot}-End.txt"), "END →\n"),
                        (os.path.join(lot_folder, f"09-{lot}-Spec.txt"), f"{spec}\n"),
                    ]

                    for file_path, content in files:
                        with open(file_path, "w", encoding="utf-8-sig") as f:
                            f.write(content)

                self.append_log(f"✔ Lot {lot} — selesai diproses")

            self.append_log("Processing complete!")
            QMessageBox.information(self, "Success", "Processing complete!")

        except Exception as e:
            self.append_log(f"Error: {e}")

    # SELECT / SEARCH
    def select_all(self):
        for row in range(self.table_widget.rowCount()):
            self.table_widget.selectRow(row)

    def deselect_all(self):
        self.table_widget.clearSelection()

    def search_table(self):
        term = self.search_box.text().lower()
        for row in range(self.table_widget.rowCount()):
            item = self.table_widget.item(row, 0)
            if item:
                self.table_widget.setRowHidden(row, term not in item.text().lower())
            else:
                self.table_widget.setRowHidden(row, False)


# ==========================================================
# RUN APP
# ==========================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelToTxtApp()
    window.show()        # <-- tetap boleh ada, tidak masalah
    sys.exit(app.exec_())
