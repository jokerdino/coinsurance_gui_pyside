from PySide6.QtCore import Qt

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QCheckBox,
    QInputDialog,
    QListWidget,
    QListWidgetItem,
)


import sys
import pandas as pd

from company_wise_reports import merge_files
from pivot_table import generate_pivot_table
from split_follower_office_code import split_files
from merge_coinsurance_files import merge_multiple_files
from premium_receivable_script import premium_receivable_function

df_premium_file = pd.DataFrame()
df_claim_file = pd.DataFrame()
df_claim_data_file = pd.DataFrame()


class CoinsuranceGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Coinsurance GUI - PySide6")
        self.setGeometry(100, 100, 600, 500)
        self.setFixedSize(575, 600)

        # Central widget
        central = QWidget()
        layout = QVBoxLayout()
        central.setLayout(layout)
        self.setCentralWidget(central)

        # Menu
        menubar = self.menuBar()
        file_menu = menubar.addMenu("File")

        about_action = file_menu.addAction("About")
        about_action.triggered.connect(self.show_version)

        quit_action = file_menu.addAction("Quit")
        quit_action.triggered.connect(self.close)

        # Buttons
        button_layout = QHBoxLayout()
        self.premium_button = QPushButton("Premium payable")
        self.premium_button.clicked.connect(self.open_premium_file)

        self.claim_button = QPushButton("Claim receivable")
        self.claim_button.clicked.connect(self.open_claim_file)

        self.claim_data_button = QPushButton("Claims data")
        self.claim_data_button.clicked.connect(self.open_multiple_claim_data_files)

        button_layout.addWidget(self.premium_button)
        button_layout.addWidget(self.claim_button)
        button_layout.addWidget(self.claim_data_button)
        layout.addLayout(button_layout)

        # Checkbox
        self.bool_hub = QCheckBox("Only OO entries")
        # layout.addWidget(self.bool_hub, alignment=Qt.AlignCenter)
        row = QHBoxLayout()
        row.addStretch()  # push left
        row.addWidget(self.bool_hub)  # place checkbox
        row.addStretch()  # push right
        layout.addLayout(row)

        self.select_all_button = QPushButton("Select all companies")
        self.select_all_button.clicked.connect(self.toggle_select_all)
        layout.addWidget(self.select_all_button)

        # Multi-select company list
        self.company_list = QListWidget()
        # self.company_list.setSelectionMode(QListWidget.MultiSelection)  # allow multiple
        layout.addWidget(self.company_list)

        button_layout_second_row = QHBoxLayout()
        # Reports buttons
        self.generate_reports_button = QPushButton("Generate reports")
        self.generate_reports_button.clicked.connect(self.generate_reports)

        row_second_row = QHBoxLayout()
        row_second_row.addStretch()  # push left
        row_second_row.addWidget(self.generate_reports_button)  # place checkbox
        row_second_row.addStretch()  # push right
        layout.addLayout(row_second_row)
        layout.addWidget(self.generate_reports_button)

        self.new_summary_button = QPushButton("Generate new summary")
        self.new_summary_button.clicked.connect(self.generate_summary)
        button_layout_second_row.addWidget(self.new_summary_button)

        self.split_button = QPushButton("Split follower office code")
        self.split_button.clicked.connect(self.split_follower_office_code)
        button_layout_second_row.addWidget(self.split_button)

        self.merge_button = QPushButton("Merge files")
        self.merge_button.clicked.connect(self.merge_multiple_files_button)
        button_layout_second_row.addWidget(self.merge_button)

        layout.addLayout(button_layout_second_row)
        self.premium_receivable_button = QPushButton("Premium receivable")
        self.premium_receivable_button.clicked.connect(self.premium_receivable)
        layout.addWidget(self.premium_receivable_button)
        self.company_list.setStyleSheet("""
            QListWidget {
                background-color: #f9f9f9;
                border: 1px solid #ccc;
                padding: 5px;
                font-size: 14px;
            }
            QListWidget::item {
                padding: 6px;
                margin: 2px 0;
            }
            QListWidget::item:selected {
                background: #d0eaff;
                color: black;
            }
            QCheckBox {
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
            }
        """)

    # ---- Methods ----
    def show_version(self):
        QMessageBox.information(self, "Version", "Current version: Qt_0.1")

    def open_premium_file(self):
        global df_premium_file
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Premium File", "", "CSV Files (*.csv)"
        )
        if file:
            df_premium_file = pd.read_csv(
                file, converters={"TXT_FOLLOWER_OFF_CD_CODE": str}
            )
            # clean company names
            df_premium_file["COMPANYNAME"] = (
                df_premium_file["COMPANYNAME"].str.replace(".", "").str.strip()
            )
            # self.update_company_dropdown()
            self.update_company_list()
            QMessageBox.information(
                self, "Message", f"Premium file {file} has been selected."
            )

    def open_claim_file(self):
        global df_claim_file
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Claim Files", "", "CSV Files (*.csv)"
        )
        if files:
            dfs = [
                pd.read_csv(
                    f,
                    converters={
                        "TXT_LEADER_OFFICE_CODE": str,
                        "TXT_FOLLOWER_OFFICE_CODE": str,
                    },
                )
                for f in files
            ]
            df_claim_file = pd.concat(dfs)
            # clean company names
            df_claim_file["COMPANYNAME"] = (
                df_claim_file["COMPANYNAME"].str.replace(".", "").str.strip()
            )
            # self.update_company_dropdown()
            self.update_company_list()
            QMessageBox.information(
                self, "Message", f"Claim files selected: {len(files)}"
            )

    def open_multiple_claim_data_files(self):
        global df_claim_data_file
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Claims Data Files", "", "CSV Files (*.csv)"
        )
        if files:
            dfs = [pd.read_csv(f, converters={"Office Code": str}) for f in files]
            df_claim_data_file = pd.concat(dfs)
            QMessageBox.information(
                self, "Message", f"Claim data files selected: {len(files)}"
            )

    def update_company_list(self):
        # collect companies from both files
        companies = set()
        if not df_premium_file.empty:
            companies.update(df_premium_file["COMPANYNAME"].unique())
        if not df_claim_file.empty:
            companies.update(df_claim_file["COMPANYNAME"].unique())

        # reset list
        self.company_list.clear()
        for c in sorted(companies):
            item = QListWidgetItem(c)

            item.setCheckState(Qt.Unchecked)  # optional: show checkboxes
            self.company_list.addItem(item)

    def toggle_select_all(self):
        all_checked = True
        for i in range(self.company_list.count()):
            item = self.company_list.item(i)
            if item.checkState() != Qt.Checked:
                all_checked = False
                break

        # If everything was checked, uncheck all; otherwise check all
        for i in range(self.company_list.count()):
            item = self.company_list.item(i)
            item.setCheckState(Qt.Unchecked if all_checked else Qt.Checked)

        # Update button text dynamically
        if all_checked:
            self.select_all_button.setText("Select All")
        else:
            self.select_all_button.setText("Deselect All")

    def generate_reports(self):
        selected_companies = []
        for i in range(self.company_list.count()):
            item = self.company_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_companies.append(item.text())

        if not selected_companies:
            QMessageBox.warning(self, "Warning", "Please select at least one company!")
            return

        path_string_wip, ok = QInputDialog.getText(
            self, "Enter folder name", "Enter folder name"
        )
        if ok and path_string_wip:
            # filter premium & claims for only selected companies
            df_premium_selected = df_premium_file[
                df_premium_file["COMPANYNAME"].isin(selected_companies)
            ]
            df_claim_selected = df_claim_file[
                df_claim_file["COMPANYNAME"].isin(selected_companies)
            ]

            merge_files(
                df_premium_selected,
                df_claim_selected,
                df_claim_data_file,
                path_string_wip,
                self.bool_hub.isChecked(),
            )

            QMessageBox.information(
                self,
                "Message",
                f"Reports generated for: {', '.join(selected_companies)}",
            )

    def generate_summary(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Company File", "", "Excel Files (*.xlsx)"
        )
        if file:
            company_name = file.split(".", 1)[0]
            generate_pivot_table(file, company_name, None)
            QMessageBox.information(
                self, "Message", f"Summary generated for {company_name}."
            )

    def split_follower_office_code(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Company File", "", "Excel Files (*.xlsx)"
        )
        if file:
            company_name = file.split(".", 1)[0]
            code, ok = QInputDialog.getText(
                self, "Enter follower office code", "Enter follower office code"
            )
            if ok:
                split_files(file, code)
                QMessageBox.information(
                    self, "Message", f"Statement split for {company_name}: {code}"
                )

    def merge_multiple_files_button(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", "", "Excel Files (*.xlsx)"
        )
        if files:
            new_file, ok = QInputDialog.getText(
                self, "Enter new file name", "Enter new file name"
            )
            if ok:
                merge_multiple_files(files, new_file.split(".", 1)[0])
                QMessageBox.information(
                    self, "Message", f"{new_file}.xlsx has been created."
                )

    def premium_receivable(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Premium Receivable CSV", "", "CSV Files (*.csv)"
        )
        if file:
            folder, ok = QInputDialog.getText(
                self, "Enter folder name", "Enter folder name for premium receivable"
            )
            if ok:
                df = pd.read_csv(file)
                premium_receivable_function(df, folder)
                QMessageBox.information(
                    self,
                    "Message",
                    f"Premium receivable generated in {folder}_PR folder.",
                )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CoinsuranceGUI()
    window.show()
    sys.exit(app.exec())
