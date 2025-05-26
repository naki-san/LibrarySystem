import sqlite3
import pandas as pd
import re
import random
import traceback
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTableWidget, QTableWidgetItem,
    QMessageBox, QFileDialog, QComboBox, QFormLayout, QHeaderView,
    QDialog, QGridLayout, QFrame, QStackedWidget, QDesktopWidget,
    QAction, QMenu
)
from PyQt5.QtGui import (
    QFont, QColor, QPixmap, QPainter, QIcon, QPalette, QBrush
)
import os
import sys
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtChart import (
    QChart, QChartView, QPieSeries, QBarSet, QBarSeries, QBarCategoryAxis, QValueAxis, QPieSlice
)
from datetime import datetime

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

connector = sqlite3.connect("library.db")
cursor = connector.cursor()

connector.execute(
    'CREATE TABLE IF NOT EXISTS Library (BK_NAME TEXT, BK_ID TEXT PRIMARY KEY NOT NULL, AUTHOR_NAME TEXT, YEAR_PUBLISHED INTEGER, CATEGORY TEXT, TOTAL_COPIES INTEGER, AVAILABLE_COPIES INTEGER, BK_STATUS TEXT)'
)

cursor.execute(
    'CREATE TABLE IF NOT EXISTS Borrowers (BORROWER_ID INTEGER PRIMARY KEY AUTOINCREMENT, BK_ID TEXT NOT NULL, BORROWER_NAME TEXT, CONTACT_NUMBER TEXT, EMAIL TEXT, GENDER TEXT, CLASSIFICATION TEXT, DATE_BORROWED TEXT, DATE_RETURNED TEXT, FOREIGN KEY (BK_ID) REFERENCES Library (BK_ID))'
)

class SidebarButton(QPushButton):
    def __init__(self, text, icon_path=None):
        super().__init__(text)
        self.setMinimumHeight(50)
        self.setCursor(Qt.PointingHandCursor)
        
        if icon_path:
            self.setIcon(QIcon(icon_path))
            self.setIconSize(QSize(24, 24))
            
        self.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #E0E0E0;
                border: none;
                text-align: left;
                padding-left: 15px;
                font-size: 16px;
                font-weight: normal;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.1);
            }
            QPushButton:checked {
                background-color: rgba(255, 255, 255, 0.2);
                border-left: 4px solid #64B5F6;
                font-weight: bold;
            }
        """)
        self.setCheckable(True)

class DashboardWidget(QWidget):
    def __init__(self, cursor, connector):
        super().__init__()
        self.cursor = cursor
        self.connector = connector
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)
        self.initUI()

    def initUI(self):
        image_path = resource_path("images/librarybackground.jpg")
        pixmap = QPixmap(image_path)
        scaled_pixmap = pixmap.scaled(self.size(), Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)

        palette = QPalette()
        palette.setBrush(self.backgroundRole(), QBrush(scaled_pixmap))
        self.setPalette(palette)
        self.setAutoFillBackground(True)
        
        self.dashboard_title = QLabel("Dashboard Overview")
        self.dashboard_title.setFont(QFont("Helvetica", 18, QFont.Bold))
        self.dashboard_title.setStyleSheet("color: white; background-color: rgba(0, 0, 0, 80); padding: 10px; border-radius: 5px;")
        self.layout.addWidget(self.dashboard_title)

        self.card_grid_layout = QGridLayout()
        self.layout.addLayout(self.card_grid_layout)

        self.total_books_card = self.create_card(resource_path("images/books.png"), "Total Books", 0, "#3498DB")
        self.issued_books_card = self.create_card(resource_path("images/issued.png"), "Issued Books", 0, "#E74C3C")
        self.student_borrowers_card = self.create_card(resource_path("images/students.png"), "Students", 0, "#2ECC71")
        self.faculty_borrowers_card = self.create_card(resource_path("images/faculty.png"), "Faculty", 0, "#F1C40F")
        self.researcher_borrowers_card = self.create_card(resource_path("images/researcher.png"), "REPS", 0, "#9B59B6")
        self.other_borrowers_card = self.create_card(resource_path("images/other.png"), "Other Borrowers", 0, "#E67E22")


        self.card_grid_layout.addWidget(self.total_books_card, 0, 0)
        self.card_grid_layout.addWidget(self.issued_books_card, 0, 1)
        self.card_grid_layout.addWidget(self.student_borrowers_card, 0, 2)
        self.card_grid_layout.addWidget(self.faculty_borrowers_card, 1, 0)
        self.card_grid_layout.addWidget(self.researcher_borrowers_card, 1, 1)
        self.card_grid_layout.addWidget(self.other_borrowers_card, 1, 2)


        self.graph_chart_layout = QHBoxLayout()
        self.borrowing_trends_tab = QWidget()
        self.classification_tab = QWidget()
        self.graph_chart_layout.addWidget(self.borrowing_trends_tab)
        self.graph_chart_layout.addWidget(self.classification_tab)

        self.layout.addLayout(self.graph_chart_layout)

        self.refresh_data()


    def refresh_data(self):
        self.total_books_card.findChild(QLabel, "value").setText(str(self.get_total_books()))
        self.issued_books_card.findChild(QLabel, "value").setText(str(self.get_issued_books()))
        self.student_borrowers_card.findChild(QLabel, "value").setText(str(self.get_borrowers_by_classification("Student")))
        self.faculty_borrowers_card.findChild(QLabel, "value").setText(str(self.get_borrowers_by_classification("Faculty")))
        self.researcher_borrowers_card.findChild(QLabel, "value").setText(str(self.get_borrowers_by_classification("REPS")))
        self.other_borrowers_card.findChild(QLabel, "value").setText(str(self.get_borrowers_by_classification("Other")))


        self.update_borrowing_trends_chart()
        self.update_classification_chart()

    def create_card(self, icon_path, title, value, color):
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: rgba(255, 255, 255, 0.85);
                border-radius: 10px;
                padding: 10px;
                border: 2px solid {color};
            }}
        """)

        card_layout = QHBoxLayout(card)
        self.card_grid_layout.setSpacing(15)
        self.card_grid_layout.setContentsMargins(10, 10, 10, 10)

        icon_label = QLabel()
        pixmap = QPixmap(icon_path).scaled(70, 70, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        icon_label.setPixmap(pixmap)
        icon_label.setAlignment(Qt.AlignVCenter)

        text_layout = QVBoxLayout()
        
        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignLeft)
        title_label.setFont(QFont("Helvetica", 12))
        title_label.setStyleSheet("color: #2C3E50; font-weight: bold;")

        value_label = QLabel(str(value))
        value_label.setAlignment(Qt.AlignLeft)
        value_label.setFont(QFont("Helvetica", 12, QFont.Bold))
        value_label.setStyleSheet(f"color: {color};")
        value_label.setObjectName("value")

        text_layout.addWidget(title_label)
        text_layout.addWidget(value_label)

        card_layout.addWidget(icon_label)
        card_layout.addLayout(text_layout)

        return card


    def update_borrowing_trends_chart(self):
        if hasattr(self, 'borrowing_trends_chart'):
            self.borrowing_trends_chart.removeAllSeries()
        else:
            self.borrowing_trends_chart = QChart()
            self.borrowing_trends_chart.setTitle("Borrowing Trends by Category")
            self.borrowing_trends_chart.setAnimationOptions(QChart.SeriesAnimations)
            self.borrowing_trends_chart.setBackgroundBrush(QColor("#F5F5F5"))

        series = QPieSeries()
        
        categories = self.get_borrowing_trends_by_category()

        for category, count in categories:
            slice_ = series.append(category, count)
            slice_.setLabel(f"{category} ({count})")
            slice_.setLabelVisible(True)
            slice_.setLabelFont(QFont("Helvetica", 9))
            slice_.setLabelBrush(QColor("#333333"))

            random_color = QColor(random.randint(50, 200), random.randint(50, 200), random.randint(50, 200))
            slice_.setBrush(random_color)

        series.setPieSize(0.75)
        series.setLabelsPosition(QPieSlice.LabelOutside)

        self.borrowing_trends_chart.addSeries(series)
        self.borrowing_trends_chart.setTitleFont(QFont("Helvetica", 12, QFont.Bold))
        self.borrowing_trends_chart.setTitleBrush(QColor("#222222"))

        if not hasattr(self, 'borrowing_chart_view'):
            self.borrowing_chart_view = QChartView(self.borrowing_trends_chart)
            self.borrowing_chart_view.setFixedSize(400, 300)
            self.borrowing_chart_view.setRenderHint(QPainter.Antialiasing)

            layout = self.borrowing_trends_tab.layout()
            if layout is None:
                layout = QHBoxLayout()
                self.borrowing_trends_tab.setLayout(layout)

            layout.addWidget(self.borrowing_chart_view)
        else:
            self.borrowing_chart_view.setChart(self.borrowing_trends_chart)


    def update_classification_chart(self):

        if hasattr(self, 'classification_chart'):
            self.classification_chart.removeAllSeries()
        else:
            self.classification_chart = QChart()
            self.classification_chart.setTitle("Borrowing Trends by Classification")
            self.classification_chart.setAnimationOptions(QChart.SeriesAnimations)
            self.classification_chart.setBackgroundBrush(QColor("#F5F5F5"))  # Light gray background

        series = QBarSeries()
        bar_set = QBarSet("Borrowers")
        categories = []
        max_value = 0

        for classification, count in self.get_borrowing_trends_by_classification():
            bar_set.append(count)
            categories.append(classification)
            max_value = max(max_value, count)

        series.append(bar_set)

        bar_set.setColor(QColor("#1E88E5"))
        bar_set.setBorderColor(QColor("#222222"))
        bar_set.setLabelFont(QFont("Helvetica", 9))
        bar_set.setLabelColor(QColor("#000000"))

        self.classification_chart.addSeries(series)

        for axis in self.classification_chart.axes():
            self.classification_chart.removeAxis(axis)

        axis_x = QBarCategoryAxis()
        axis_x.append(categories)
        axis_x.setGridLineColor(QColor("#666666"))
        self.classification_chart.addAxis(axis_x, Qt.AlignBottom)
        series.attachAxis(axis_x)

        axis_y = QValueAxis()
        axis_y.setRange(0, max_value)  
        axis_y.setTickCount(max_value + 1)  
        axis_y.setMinorTickCount(0)  
        axis_y.setLabelFormat("%d")  
        axis_y.setGridLineColor(QColor("#666666"))
        self.classification_chart.addAxis(axis_y, Qt.AlignLeft)
        series.attachAxis(axis_y)

        self.classification_chart.setTitleFont(QFont("Helvetica", 12, QFont.Bold))
        self.classification_chart.setTitleBrush(QColor("#222222"))

        if not hasattr(self, 'classification_chart_view'):
            self.classification_chart_view = QChartView(self.classification_chart)
            self.classification_chart_view.setFixedSize(500, 300) 
            self.classification_chart_view.setRenderHint(QPainter.Antialiasing)

            layout = self.classification_tab.layout()
            if layout is None:
                layout = QHBoxLayout()
                self.classification_tab.setLayout(layout)

            layout.addWidget(self.classification_chart_view)
        else:
            self.classification_chart_view.setChart(self.classification_chart)

    def get_total_books(self):
        self.cursor.execute("SELECT SUM(TOTAL_COPIES) FROM Library")
        return self.cursor.fetchone()[0]

    def get_issued_books(self):
        self.cursor.execute("SELECT COUNT(*) FROM Borrowers WHERE DATE_RETURNED IS NULL")
        return self.cursor.fetchone()[0]

    def get_borrowers_by_classification(self, classification):
        self.cursor.execute("""
            SELECT COUNT(*) FROM Borrowers 
            WHERE CLASSIFICATION = ? AND DATE_RETURNED IS NULL
        """, (classification,))
        return self.cursor.fetchone()[0]

    def get_other_borrowers_count(self):
        self.cursor.execute("""
            SELECT COUNT(*) FROM Borrowers 
            WHERE CLASSIFICATION NOT IN ('Student', 'Faculty', 'REPS')
        """)
        result = self.cursor.fetchone()
        return result[0] if result else 0

    
    def get_borrowing_trends_by_category(self):
        self.cursor.execute("""
            SELECT Library.CATEGORY, COUNT(*) as count 
            FROM Borrowers 
            JOIN Library ON Borrowers.BK_ID = Library.BK_ID 
            WHERE Borrowers.DATE_RETURNED IS NULL 
            GROUP BY Library.CATEGORY
        """)
        return self.cursor.fetchall()


    def get_borrowing_trends_by_classification(self):
        self.cursor.execute("""
            SELECT CLASSIFICATION, COUNT(*) 
            FROM Borrowers 
            WHERE DATE_RETURNED IS NULL 
            GROUP BY CLASSIFICATION
        """)
        return self.cursor.fetchall()


class InventoryWidget(QWidget):
    def __init__(self, main_window, cursor, connector):
        super().__init__()
        self.main_window = main_window
        self.cursor = cursor
        self.connector = connector
        self.setMaximumSize(1920, 1080)
        self.initUI()

    def initUI(self):

        image_path = resource_path("images/librarybg.jpg")
        pixmap = QPixmap(image_path)
        scaled_pixmap = pixmap.scaled(self.size(), Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)

        palette = QPalette()
        palette.setBrush(QPalette.Background, QBrush(scaled_pixmap))
        self.setPalette(palette)
        self.setAutoFillBackground(True)
    
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        inventory_title = QLabel("CISC Library Inventory System")
        inventory_title.setFont(QFont("Helvetica", 18, QFont.Bold))
        inventory_title.setStyleSheet("color: white; background-color: rgba(0, 0, 0, 80); padding: 10px; border-radius: 5px;")
        layout.addWidget(inventory_title)

        self.top_layout = QHBoxLayout()
        
        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText("Search by Name, Author, or ID")
        self.search_input.setStyleSheet("""
            QLineEdit {
                border: 2px solid #3498DB;
                border-radius: 5px;
                padding: 8px;
                font-size: 14px;
            }
        """)
        self.top_layout.addWidget(self.search_input)
        

        button_base_color = "#3498DB"     
        button_hover_color = "#2E86C1"    

        self.search_button = QPushButton("🔍 Search", self)
        self.search_input.textChanged.connect(self.search_record)
        self.search_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {button_base_color};
                color: white;
                font-size: 14px;
                border: none;
                border-radius: 5px;
                padding: 10px 15px;
            }}
            QPushButton:hover {{
                background-color: {button_hover_color};
            }}
        """)
        self.top_layout.addWidget(self.search_button)

        self.import_button = QPushButton("📥 Import from Excel", self)
        self.import_button.clicked.connect(self.import_from_excel)
        self.import_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {button_base_color};
                color: white;
                font-size: 14px;
                border: none;
                border-radius: 5px;
                padding: 10px 15px;
            }}
            QPushButton:hover {{
                background-color: {button_hover_color};
            }}
        """)
        self.top_layout.addWidget(self.import_button)

        
        layout.addLayout(self.top_layout)
        
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(["Book Title", "Book ID", "Author", "Year", "Category", "Total\nCopies", "Available\nCopies", "Status"])

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_book_context_menu)

        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: 2px solid #3498DB;
                border-radius: 8px;
                gridline-color: #D6DBDF;
                selection-background-color: #85C1E9;
                font-size: 14px;
                color: #2C3E50;
            }

            QHeaderView::section {
                background-color: #2980B9;
                color: white;
                font-weight: bold;
                padding: 8px;
                border: none;
                font-size: 14px;
            }

            QTableWidget::item {
                padding: 8px;
                border-radius: 4px;
            }

            QTableWidget::item:selected {
                background-color: #AED6F1;
                color: black;
            }

            QTableWidget::item:hover {
                background-color: #D4E6F1;
            }

            QTableCornerButton::section {
                background-color: #2980B9;
                border-radius: 5px;
            }
        """)


        layout.addWidget(self.table)
        
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)  
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents) 
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)  
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents) 
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch) 
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents) 
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents) 
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)

        self.button_layout = QHBoxLayout()

        blue_base_color = "#3498DB"
        blue_hover_color = "#2E86C1"

        delete_hover_color = "#C0392B"

        self.add_button = self.create_action_button("➕ Add Record", self.add_record, blue_base_color, blue_hover_color)
        self.update_button = self.create_action_button("✏️ Update Record", self.update_record, blue_base_color, blue_hover_color)
        self.borrow_button = self.create_action_button("📚 Borrow Book", self.borrow_book, blue_base_color, blue_hover_color)
        self.clear_button = self.create_action_button("🔄 Clear Fields", self.clear_fields, blue_base_color, blue_hover_color)

        self.delete_button = self.create_action_button("❌ Delete Record", self.remove_record, blue_base_color, delete_hover_color)
        self.delete_all_button = self.create_action_button("🗑️ Delete All", self.remove_all_records, blue_base_color, delete_hover_color)


        button_grid = QGridLayout()
        button_grid.addWidget(self.add_button, 0, 0)
        button_grid.addWidget(self.update_button, 0, 1)
        button_grid.addWidget(self.borrow_button, 0, 2)
        button_grid.addWidget(self.delete_button, 0, 3)
        button_grid.addWidget(self.delete_all_button, 1, 1)
        button_grid.addWidget(self.clear_button, 1, 2)
        
        layout.addLayout(button_grid)
        
        self.load_records()

    def create_action_button(self, text, callback, base_color, hover_color):
        button = QPushButton(text)
        button.clicked.connect(callback)
        button.setStyleSheet(f"""
            QPushButton {{
                background-color: {base_color};
                color: white;
                font-size: 14px;
                border: none;
                border-radius: 5px;
                padding: 10px 15px;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
        """)
        return button


    def load_records(self):
        self.table.setRowCount(0)
        cursor.execute("SELECT * FROM Library")
        records = cursor.fetchall() or []

        for row_num, row_data in enumerate(records):
            self.table.insertRow(row_num)
            for col_num, data in enumerate(row_data):
                item = QTableWidgetItem(str(data))
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

                item.setToolTip(str(data))  

                self.table.setItem(row_num, col_num, item)

    def show_book_context_menu(self, position):
        menu = QMenu()

        view_details_action = QAction("📖 View Full Book Details", self)
        view_details_action.triggered.connect(self.view_full_book_details)

        menu.addAction(view_details_action)
        menu.exec_(self.table.viewport().mapToGlobal(position))


    def view_full_book_details(self):
        selected_row = self.table.currentRow()

        if selected_row == -1:
            QMessageBox.warning(self, "Error", "Please select a book to view details!")
            return

        book_id = self.table.item(selected_row, 0).text().strip()
        book_title = self.table.item(selected_row, 1).text().strip()
        author = self.table.item(selected_row, 2).text().strip()
        year = self.table.item(selected_row, 3).text().strip()
        category = self.table.item(selected_row, 4).text().strip()
        total_copies = self.table.item(selected_row, 5).text().strip()
        available_copies = self.table.item(selected_row, 6).text().strip()
        status = self.table.item(selected_row, 7).text().strip()

        import re
        formatted_title = re.sub(r'(\d+\.)', r'\1 ', book_title).strip()

        details_text = f"""
        <html>
        <head>
        <style>
            body {{
                font-family: Helvetica, Arial, sans-serif;
                font-size: 14px;
                color: #2C3E50;
            }}
            b {{
                font-weight: bold;
                color: black;
            }}
            p {{
                margin: 2px 0;
                line-height: 1.4;
            }}
        </style>
        </head>
        <body>
        <p><b>Book ID:</b> {book_id}</p>
        <p><b>Book Title:</b> {formatted_title}</p>
        <p><b>Author:</b> {author}</p>
        <p><b>Year:</b> {year}</p>
        <p><b>Category:</b> {category}</p>
        <p><b>Total Copies:</b> {total_copies}</p>
        <p><b>Available Copies:</b> {available_copies}</p>
        <p><b>Status:</b> {status}</p>
        </body>
        </html>
        """

        self.book_details_window = QDialog(self)
        self.book_details_window.setWindowTitle("Book Details")
        self.book_details_window.setGeometry(450, 250, 500, 320)
        self.book_details_window.setStyleSheet("background-color: white;")

        layout = QVBoxLayout()

        details_label = QLabel(details_text)
        details_label.setWordWrap(True)
        details_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        details_label.setStyleSheet("padding: 10px; font-size: 14px;")

        ok_button = QPushButton("OK")
        ok_button.setStyleSheet("""
            background-color: #3498DB; 
            color: white; 
            font-size: 14px;
            font-weight: bold;
            padding: 8px;
            min-width: 100px;
            border-radius: 5px;
        """)
        ok_button.clicked.connect(self.book_details_window.close)

        layout.addWidget(details_label)
        layout.addWidget(ok_button, alignment=Qt.AlignCenter)
        self.book_details_window.setLayout(layout)

        self.book_details_window.exec_()


    def search_record(self):
        query = self.search_input.text().strip()
        self.table.setRowCount(0)

        if not query:
            cursor.execute("SELECT * FROM Library")
        else:
            cursor.execute("""
                SELECT * FROM Library 
                WHERE BK_NAME LIKE ? OR AUTHOR_NAME LIKE ? OR BK_ID LIKE ?
            """, (f'%{query}%', f'%{query}%', f'%{query}%'))

        results = cursor.fetchall()

        if results:
            for row_num, row_data in enumerate(results):
                self.table.insertRow(row_num)
                for col_num, data in enumerate(row_data):
                    item = QTableWidgetItem(str(data))

                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

                    item.setToolTip(str(data))

                    self.table.setItem(row_num, col_num, item)
        else:
            if query:
                self.table.setRowCount(1)
                item = QTableWidgetItem("No matching records found")
                item.setFlags(Qt.ItemIsEnabled)
                self.table.setItem(0, 0, item)
                self.table.setSpan(0, 0, 1, self.table.columnCount())


    def import_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return

        try:
            xls = pd.ExcelFile(file_path)
            excluded_sheets = ["Categories_Key"]
            imported_books = 0  
            skipped_books = 0  
            existing_book_ids = set() 
            unknown_counter = 1 
            
            cursor.execute("SELECT BK_ID FROM Library")
            for row in cursor.fetchall():
                existing_book_ids.add(row[0])

            category_mappings = {}
            
            metadata_sheet = None
            for sheet in xls.sheet_names:
                if sheet in excluded_sheets or "Categories" in sheet or "Key" in sheet:
                    metadata_sheet = sheet
                    break
            
            if metadata_sheet:
                for header_row in range(0, 10):
                    try:
                        metadata_df = pd.read_excel(file_path, sheet_name=metadata_sheet, header=header_row)
                        
                        code_col = None
                        title_col = None
                        
                        for col in metadata_df.columns:
                            col_str = str(col).upper()
                            if any(term in col_str for term in ["CODE", "ABBREV", "ABBREVIATION", "ID"]):
                                code_col = col
                            elif "TITLE" in col_str or "NAME" in col_str or "DESCRIPTION" in col_str:
                                title_col = col
                        
                        if code_col and title_col:
                            for _, row in metadata_df.iterrows():
                                code = str(row[code_col]).strip()
                                title = str(row[title_col]).strip()
                                if code and title and code != "nan" and title != "nan":
                                    category_mappings[code] = title
                            break
                    except Exception as e:
                        print(f"Error reading metadata sheet with header row {header_row}: {e}")
            
            if not category_mappings:
                default_categories = {
                    "AB": "ANNOTATED BIBLIOGRAPHY",
                    "ADG": "ARTICLES ON DATA GATHERING",
                    "AF": "ARTICLES ON FORESTRY",
                    "ALR": "ARTICLES ON LAND/AGRARIAN REFORM",
                    "ANREP_CISC": "ANNUAL REPORTS_CISC",
                    "ANREP_OTHER": "ANNUAL REPORTS",
                    "AR": "READING MATERIALS ON AGRARIAN REFORM",
                    "ARCCESS_ND": "ARCCESS PROJECTS",
                    "ARCCESS_OEND": "OE NADA",
                    "ASD": "ARTICLES ON SUSTAINABLE DEVELOPMENT",
                    "B": "BOOKS",
                    "BD": "ASIAN BIOTECHNOLOGY AND DEVELOPMENT REVIEW",
                    "C": "Census",
                    "CARP": "READING MATERIALS ON COMPREHENSIVE AGRARIAN REFORM PROGRAM (CARP)",
                    "CDS": "CONFERENCE/DIALOGUES/SYMPOSIUM/SEMINAR",
                    "CPB": "CPAF POLICY BRIEF",
                    "DFO": "DEVELOPMENT/FRAMEWORK/OPERATIONAL PLAN",
                    "DP": "DISCUSSION PAPER SERIES",
                    "FS": "READING MATERIALS ON FORESTRY",
                    "IDP": "IARDS/CPAF DEVELOPMENT PLAN",
                    "ISF": "RESEARCH STUDIES ON INTEGRATED SOCIAL FORESTRY (ISF) AREAS",
                    "J": "JOURNAL",
                    "M": "Manuals",
                    "MP": "MASTER PLAN",
                    "MS": "MONOGRAPH SERIES",
                    "OP": "OCCASIONAL PAPER",
                    "P": "PROCEEDINGS",
                    "PAM": "PAMPHLETS",
                }
                category_mappings.update(default_categories)

            for sheet_name in xls.sheet_names:
                if sheet_name in excluded_sheets or sheet_name == metadata_sheet:
                    continue

                print(f"Processing sheet: {sheet_name}")
                
                max_header_row = 15  
                df = None
                header_row = None
                
                header_indicators = ["title", "author", "publisher", "no.", "id", "year", "copies"]
                
                for i in range(max_header_row):
                    try:
                        temp_df = pd.read_excel(file_path, sheet_name=sheet_name, header=i)
                        
                        header_matches = 0
                        for col in temp_df.columns:
                            col_str = str(col).lower()
                            if any(indicator in col_str for indicator in header_indicators):
                                header_matches += 1
                        
                        if header_matches >= 2:
                            df = temp_df
                            header_row = i
                            print(f"Found header row at index {i} with {header_matches} matches")
                            break
                    except Exception as e:
                        print(f"Error checking row {i}: {e}")

                if df is None:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
                    header_row = 0
                    print("Using first row as header by default")

                category_code = None
                category_name = None

                sheet_code = sheet_name.replace("Copy of ", "").strip()

                match = re.search(r"\((.*?)\)", sheet_code)
                if match:
                    parenthesis_text = match.group(1)
                    sheet_code = re.sub(r"\(.*?\)", "", sheet_code).strip()

                    if parenthesis_text.upper() == "OE-NADA":
                        short_parenthesis = "OEND"
                    elif parenthesis_text.upper() == "NADA":
                        short_parenthesis = "ND"
                    else:
                        short_parenthesis = "".join([word[:2].upper() for word in parenthesis_text.split()])
                    sheet_code = f"{sheet_code}_{short_parenthesis}"
                
                sheet_code = sheet_code.replace(" ", "_")

                if sheet_code in category_mappings:
                    category_code = sheet_code
                    category_name = category_mappings[sheet_code]
                else:
                    try:
                        for i in range(max(0, header_row-5), header_row):
                            row_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=1, skiprows=i)
                            for j in range(len(row_df.columns)):
                                cell_value = str(row_df.iloc[0, j]).strip().upper()
                                for code in category_mappings.keys():
                                    if code == cell_value:
                                        category_code = code
                                        category_name = category_mappings[code]
                                        break
                            if category_code:
                                break
                    except Exception as e:
                        print(f"Error looking for category in header area: {e}")
                
                if not category_code:
                    category_code = sheet_code
                    category_name = sheet_code
                    
                    for code, name in category_mappings.items():
                        if code in sheet_name.upper().replace(" ", "_"):
                            category_code = code
                            category_name = name
                            break
                
                print(f"Using category: {category_code} - {category_name}")
                
                book_id_col = None
                title_col = None
                author_col = None
                year_col = None
                copies_col = None
                
                for col in df.columns:
                    col_str = str(col).lower()
                    
                    if category_code and category_code.lower() in col_str and any(x in col_str for x in ["no", "number", "id"]):
                        book_id_col = col
                    elif "no." in col_str and "copies" not in col_str:
                        book_id_col = col
                    
                    if "title" in col_str:
                        title_col = col
                    elif any(term in col_str for term in ["author", "publisher", "compiler"]):
                        author_col = col
                    elif any(term in col_str for term in ["year", "date", "published"]):
                        year_col = col
                    elif any(term in col_str for term in ["copies", "copy", "quantity"]):
                        copies_col = col

                if not book_id_col:
                    for col in df.columns:
                        col_str = str(col).lower()
                        if any(term in col_str for term in ["id", "number", "code", "call no"]):
                            book_id_col = col
                            break
                
                if not book_id_col:
                    QMessageBox.warning(self, "Error", f"Could not find Book ID column in {sheet_name}. Skipping...")
                    continue
                
                column_mapping = {}
                if title_col:
                    column_mapping[title_col] = "BK_NAME"
                if book_id_col:
                    column_mapping[book_id_col] = "BK_ID"
                if author_col:
                    column_mapping[author_col] = "AUTHOR_NAME"
                if year_col:
                    column_mapping[year_col] = "YEAR_PUBLISHED"
                if copies_col:
                    column_mapping[copies_col] = "TOTAL_COPIES"
                
                df = df.rename(columns=column_mapping)
                
                required_columns = ["BK_NAME", "BK_ID", "AUTHOR_NAME", "YEAR_PUBLISHED", "TOTAL_COPIES"]
                for col in required_columns:
                    if col not in df.columns:
                        df[col] = "-"
                
                for col in df.columns:
                    if df[col].dtype == "object":
                        df.loc[:, col] = df[col].fillna("-")
                
                df["TOTAL_COPIES"] = df["TOTAL_COPIES"].replace("NO COPIES FOUND", 0)
                df["TOTAL_COPIES"] = pd.to_numeric(df["TOTAL_COPIES"], errors='coerce').fillna(0).astype(int)
                df["YEAR_PUBLISHED"] = pd.to_numeric(df["YEAR_PUBLISHED"], errors='coerce').fillna(0).astype(int)

                df["BK_ID"] = df["BK_ID"].astype(str).str.replace(r"\.0$", "", regex=True)

                def generate_book_id(row):
                    nonlocal unknown_counter
                    
                    original_id = str(row["BK_ID"]).strip()

                    original_id = re.sub(r"[^a-zA-Z0-9_-]", "", original_id)

                    if original_id in ["-", "", "nan"] or original_id.isspace():
                        while True:
                            new_id = f"UNKNOWN_{unknown_counter}"
                            unknown_counter += 1
                            
                            if new_id not in existing_book_ids:
                                break
                    else:
                        numeric_part = re.search(r'\d+', original_id)
                        if numeric_part:
                            numeric_id = numeric_part.group()
                            new_id = f"{category_code}{numeric_id}"
                        else:
                            new_id = f"{category_code}{original_id}"

                    counter = 1
                    unique_id = new_id
                    while unique_id in existing_book_ids:
                        unique_id = f"{new_id}_{counter}"
                        counter += 1
                    
                    existing_book_ids.add(unique_id)
                    return unique_id

                df["BK_ID"] = df.apply(generate_book_id, axis=1)
                

                df["AVAILABLE_COPIES"] = df["TOTAL_COPIES"]
                df["BK_STATUS"] = df["AVAILABLE_COPIES"].apply(lambda x: "Available" if x > 0 else "Fully Issued")

                df["CATEGORY"] = category_name

                df = df[df["BK_NAME"].str.lower() != "title"]

                for _, row in df.iterrows():
                    book_name = str(row["BK_NAME"]).strip()

                    if book_name in ["-", "", "nan"] or not book_name:
                        book_name = f"[Untitled Book {row['BK_ID']}]"
                        skipped_books += 1 
                        
                    try:
                        cursor.execute(
                            """INSERT INTO Library (BK_NAME, BK_ID, AUTHOR_NAME, YEAR_PUBLISHED, CATEGORY,
                            TOTAL_COPIES, AVAILABLE_COPIES, BK_STATUS) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                            (book_name, row["BK_ID"], row["AUTHOR_NAME"], row["YEAR_PUBLISHED"], 
                            row["CATEGORY"], row["TOTAL_COPIES"], row["AVAILABLE_COPIES"], row["BK_STATUS"])
                        )
                        imported_books += 1
                    except sqlite3.IntegrityError:
                        QMessageBox.warning(self, "Duplicate Book ID", f"Skipping duplicate Book ID: {row['BK_ID']}")
                    except Exception as e:
                        print(f"Error inserting row: {e}")
            
            connector.commit()

            QMessageBox.information(self, "Import Results", 
                f"{imported_books} books imported successfully!\n"
                f"{skipped_books} books had no titles and were imported with placeholder names.")

            connector.commit()

            self.load_records()

            if hasattr(self.main_window, 'dashboard_widget'):
                try:
                    self.main_window.dashboard_widget.refresh_data()
                    print("Dashboard refreshed successfully")
                except Exception as e:
                    print(f"Error refreshing dashboard: {e}")
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to import Excel file: {e}")
            print(f"Exception details: {traceback.format_exc()}") 
        
        
    def add_record(self):
        self.add_window = QWidget()
        self.add_window.setWindowTitle("Add New Book")
        self.add_window.setGeometry(200, 200, 400, 300)
        layout = QVBoxLayout()
        self.add_window.setStyleSheet("""
            QWidget {
                background-color: white;
                font-family: Helvetica;
            }
            QLineEdit {
                border: 2px solid #3498DB;
                border-radius: 5px;
                padding: 8px;
                font-size: 14px;
            }
            QPushButton {
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
                color: white;
            }
            QLabel {
                font-size: 14px;
                color: #2C3E50;
            }
        """)

        form_layout = QFormLayout()
        self.book_name_input = QLineEdit()
        self.book_id_input = QLineEdit()
        self.author_input = QLineEdit()
        self.year_input = QLineEdit()
        self.category_input = QLineEdit()
        self.total_copies_input = QLineEdit()

        self.book_name_input = QLineEdit()
        self.book_name_input.setPlaceholderText("Enter book title")

        self.book_id_input = QLineEdit()
        self.book_id_input.setPlaceholderText("Enter book ID")

        self.author_input = QLineEdit()
        self.author_input.setPlaceholderText("Enter author(s) name")

        self.year_input = QLineEdit()
        self.year_input.setPlaceholderText("Enter 4-digit year or 0 if N/A") 

        self.category_input = QLineEdit()
        self.category_input.setPlaceholderText("Enter book category")

        self.total_copies_input = QLineEdit()
        self.total_copies_input.setPlaceholderText("Enter total copies")


        form_layout.addRow("Book Title:", self.book_name_input)
        form_layout.addRow("Book ID:", self.book_id_input)
        form_layout.addRow("Authors:", self.author_input)
        form_layout.addRow("Year Published:", self.year_input)
        form_layout.addRow("Category:", self.category_input)
        form_layout.addRow("Total Copies:", self.total_copies_input)


        layout.addLayout(form_layout)

        button_layout = QHBoxLayout()
        
        save_button = QPushButton("Save")
        save_button.setStyleSheet("background-color: #27AE60;")
        save_button.clicked.connect(self.save_record)
        button_layout.addWidget(save_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.setStyleSheet("background-color: #E74C3C;")
        cancel_button.clicked.connect(self.add_window.close)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.add_window.setLayout(layout)
        self.add_window.show()

    def save_record(self):
        book_name = self.book_name_input.text().strip()
        book_id = self.book_id_input.text().strip()
        author = self.author_input.text().strip()
        year_published = self.year_input.text().strip()
        category = self.category_input.text().strip()
        total_copies = self.total_copies_input.text().strip()

        if not all([book_name, book_id, author, year_published, category, total_copies]):
            QMessageBox.warning(self, "Error", "All fields are required!\n\nNote: If no available year, enter 0.")
            self.add_window.raise_()
            self.add_window.activateWindow()
            return

        if not year_published.isdigit() or (year_published != "0" and len(year_published) != 4):
            QMessageBox.warning(self, "Error", "Please enter a valid 4-digit year or 0 if unavailable!")
            self.add_window.raise_()
            self.add_window.activateWindow()
            return

        if not total_copies.isdigit() or int(total_copies) < 1:
            QMessageBox.warning(self, "Error", "Total copies must be at least 1!")
            self.add_window.raise_()
            self.add_window.activateWindow()
            return

        total_copies = int(total_copies)
        available_copies = total_copies
        status = "Available" if available_copies > 0 else "Fully Issued"

        try:
            cursor.execute(
                "INSERT INTO Library (BK_NAME, BK_ID, AUTHOR_NAME, YEAR_PUBLISHED, CATEGORY, TOTAL_COPIES, AVAILABLE_COPIES, BK_STATUS) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (book_name, book_id, author, int(year_published), category, total_copies, available_copies, status)
            )
            connector.commit()
            QMessageBox.information(self, "Success", "Book added successfully!\n\nNote: If no available year, enter 0.")
            self.add_window.close()
            self.load_records()

            if hasattr(self.main_window, 'dashboard_widget'):
                self.main_window.dashboard_widget.refresh_data()

        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Error", "Book ID already exists!")
            self.add_window.raise_()
            self.add_window.activateWindow()


    def update_record(self):
            selected_row = self.table.currentRow()

            if selected_row == -1:
                QMessageBox.warning(self, "Error", "Please select a book record to update!")
                return

            book_name = self.table.item(selected_row, 0).text()
            book_id = self.table.item(selected_row, 1).text()
            author = self.table.item(selected_row, 2).text()
            year_published = self.table.item(selected_row, 3).text()
            category = self.table.item(selected_row, 4).text()
            total_copies = self.table.item(selected_row, 5).text()
            available_copies = self.table.item(selected_row, 6).text()

            self.update_window = QWidget()
            self.update_window.setWindowTitle("Update Book Record")
            self.update_window.setGeometry(300, 200, 400, 400)
            
            self.update_window.setStyleSheet("""
                QWidget {
                    background-color: white;
                    font-family: Helvetica;
                }
                QLineEdit {
                    border: 2px solid #3498DB;  /* Blue border for input fields */
                    border-radius: 5px;
                    padding: 8px;
                    font-size: 14px;
                    color: black;  /* Text color for inputs */
                }
                QPushButton {
                    border-radius: 5px;
                    padding: 10px;
                    font-size: 14px;
                    color: white;
                }
                QPushButton#updateButton {
                    background-color: #3498DB;  /* Blue for Update button */
                }
                QPushButton#updateButton:hover {
                    background-color: #2980B9;  /* Darker blue on hover */
                }
                QPushButton#cancelButton {
                    background-color: #E74C3C;  /* Red for Cancel button */
                }
                QPushButton#cancelButton:hover {
                    background-color: #C0392B;  /* Darker red on hover */
                }
                QLabel {
                    font-size: 14px;
                    color: black;  /* Set text color of labels to black */
                }
            """)

            layout = QVBoxLayout()

            form_layout = QFormLayout()
            self.update_book_name = QLineEdit(book_name)
            self.update_author = QLineEdit(author)
            self.update_year = QLineEdit(year_published)
            self.update_category = QLineEdit(category)
            self.update_total_copies = QLineEdit(total_copies)

            self.update_book_id = QLineEdit(book_id)
            self.update_book_id.setReadOnly(True)
            self.update_book_id.setStyleSheet("background-color: #f0f0f0;")

            form_layout.addRow("Book Title:", self.update_book_name)
            form_layout.addRow("Book ID (Read-only):", self.update_book_id)
            form_layout.addRow("Authors:", self.update_author)
            form_layout.addRow("Year Published:", self.update_year)
            form_layout.addRow("Category:", self.update_category)
            form_layout.addRow("Total Copies:", self.update_total_copies)

            layout.addLayout(form_layout)

            button_layout = QHBoxLayout()

            update_button = QPushButton("Update")
            update_button.setObjectName("updateButton")
            update_button.clicked.connect(lambda: self.save_updated_record(book_id, int(available_copies)))
            button_layout.addWidget(update_button)

            cancel_button = QPushButton("Cancel")
            cancel_button.setObjectName("cancelButton")
            cancel_button.clicked.connect(self.update_window.close)
            button_layout.addWidget(cancel_button)

            layout.addLayout(button_layout)
            self.update_window.setLayout(layout)
            self.update_window.show()

    def save_updated_record(self, book_id, old_available_copies):
        book_name = self.update_book_name.text().strip()
        author = self.update_author.text().strip()
        year_published = self.update_year.text().strip()
        category = self.update_category.text().strip()
        total_copies_input = self.update_total_copies.text().strip()

        if not book_name:
            QMessageBox.warning(self, "Error", "Book title is required!")
            self.update_window.raise_()
            self.update_window.activateWindow()
            return
        if not author:
            QMessageBox.warning(self, "Error", "Author name is required!")
            self.update_window.raise_()
            self.update_window.activateWindow()
            return
        if not category:
            QMessageBox.warning(self, "Error", "Category is required!")
            self.update_window.raise_()
            self.update_window.activateWindow()
            return
        if not total_copies_input.isdigit() or int(total_copies_input) < 1:
            QMessageBox.warning(self, "Error", "Total copies must be a valid positive number!")
            self.update_window.raise_()
            self.update_window.activateWindow()
            return

        if not year_published.isdigit() or (year_published != "0" and len(year_published) != 4):
            QMessageBox.warning(self, "Error", "Please enter a valid 4-digit year or 0 if unavailable!")
            self.update_window.raise_()
            self.update_window.activateWindow()
            return

        new_total_copies = int(total_copies_input)

        cursor.execute("SELECT TOTAL_COPIES, AVAILABLE_COPIES FROM Library WHERE BK_ID = ?", (book_id,))
        book_data = cursor.fetchone()

        if not book_data:
            QMessageBox.warning(self, "Error", "Book not found in the database!")
            return

        old_total_copies, available_copies = book_data
        borrowed_count = old_total_copies - available_copies

        if new_total_copies < borrowed_count:
            QMessageBox.warning(self, "Error", f"Cannot set total copies lower than borrowed copies ({borrowed_count})!")
            return

        new_available_copies = new_total_copies - borrowed_count
        new_status = "Available" if new_available_copies > 0 else "Fully Issued"

        try:
            cursor.execute(
                "UPDATE Library SET BK_NAME=?, AUTHOR_NAME=?, YEAR_PUBLISHED=?, CATEGORY=?, TOTAL_COPIES=?, AVAILABLE_COPIES=?, BK_STATUS=? WHERE BK_ID=?",  
                (book_name, author, int(year_published), category, new_total_copies, new_available_copies, new_status, book_id)
            )
            connector.commit()
            QMessageBox.information(self, "Success", "Book record updated successfully!")
            self.update_window.close()
            self.load_records()

            if hasattr(self.main_window, 'dashboard_widget'):
                self.main_window.dashboard_widget.refresh_data()

        except sqlite3.Error as e:
            QMessageBox.critical(self, "Error", f"Failed to update record: {e}")


    def borrow_book(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Error", "Please select a book to borrow!")
            return

        book_name = self.table.item(selected_row, 0).text()
        book_id = self.table.item(selected_row, 1).text()

        cursor.execute("SELECT AVAILABLE_COPIES, BK_STATUS FROM Library WHERE BK_ID = ?", (book_id,))
        book_data = cursor.fetchone()

        if not book_data:
            QMessageBox.warning(self, "Error", "Book not found in the database!")
            return

        available_copies, status = book_data

        if available_copies <= 0 or status == "Fully Issued":
            QMessageBox.warning(self, "Error", "This book is currently fully issued! No available copies.")
            return

        self.borrow_window = QWidget()
        self.borrow_window.setWindowTitle("Borrow Book")
        self.borrow_window.setFixedSize(820, 490) 
        self.borrow_window.setStyleSheet("""
            QWidget {
                background-color: #F5F5F5;
                font-family: Helvetica;
            }
            QLabel {
                font-size: 14px;
                color: #2C3E50;
            }
            QLineEdit, QComboBox {
                border: 2px solid #3498DB;
                border-radius: 5px;
                padding: 8px;
                font-size: 14px;
                background-color: white;
            }
            QPushButton {
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
                color: white;
                background-color: #3498DB;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        header_label = QLabel("Borrow Book")
        header_label.setFont(QFont("Helvetica", 18, QFont.Bold))
        header_label.setStyleSheet("color: #3498DB; margin-bottom: 15px;")
        layout.addWidget(header_label)

        book_info_layout = QFormLayout()
        book_info_layout.setSpacing(10)

        book_name_label = QLabel(book_name)
        book_name_label.setWordWrap(True) 
        book_name_label.setFixedWidth(650) 
        book_name_label.setStyleSheet("font-size: 16px; color: #2C3E50; font-weight: bold;")
        book_info_layout.addRow("Book Title:", book_name_label)

        book_id_label = QLabel(book_id)
        book_id_label.setStyleSheet("font-size: 14px; color: #2C3E50;")
        book_info_layout.addRow("Book ID:", book_id_label)

        layout.addLayout(book_info_layout)

        borrower_form_layout = QFormLayout()
        borrower_form_layout.setSpacing(10)

        self.borrower_name_input = QLineEdit()
        self.borrower_name_input.setPlaceholderText("Enter borrower's name")
        borrower_form_layout.addRow("Borrower Name:", self.borrower_name_input)

        self.contact_input = QLineEdit()
        self.contact_input.setPlaceholderText("Enter contact number")
        borrower_form_layout.addRow("Contact:", self.contact_input)

        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("Enter email address")
        borrower_form_layout.addRow("Email:", self.email_input)

        self.gender_input = QComboBox()
        self.gender_input.addItems(["Male", "Female"])
        borrower_form_layout.addRow("Gender:", self.gender_input)

        self.classification_input = QComboBox()
        self.classification_input.addItems(["Student", "Faculty", "REPS", "Other (specify)"])
        self.classification_input.currentIndexChanged.connect(self.toggle_other_input)

        self.other_classification_input = QLineEdit()
        self.other_classification_input.setPlaceholderText("Please specify classification")
        self.other_classification_input.setFixedHeight(0)  
        self.other_classification_input.setVisible(False) 

        classification_layout = QVBoxLayout()
        classification_layout.addWidget(self.classification_input)
        classification_layout.addWidget(self.other_classification_input)

        borrower_form_layout.addRow("Classification:", classification_layout)

        layout.addLayout(borrower_form_layout)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)

        save_button = QPushButton("📚 Borrow Book")
        save_button.setStyleSheet("background-color: #27AE60;")
        save_button.clicked.connect(lambda: self.confirm_borrow(book_id))
        button_layout.addWidget(save_button)

        cancel_button = QPushButton("🚫 Cancel")
        cancel_button.setStyleSheet("background-color: #E74C3C;")
        cancel_button.clicked.connect(self.borrow_window.close)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        self.borrow_window.setLayout(layout)
        self.borrow_window.show()

    def toggle_other_input(self):
        selected = self.classification_input.currentText()
        if selected == "Other (specify)":
            self.other_classification_input.setVisible(True)
            self.other_classification_input.setFixedHeight(40) 
        else:
            self.other_classification_input.setVisible(False)
            self.other_classification_input.setFixedHeight(0)


    def is_valid_email(self, email):
        pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(pattern, email) is not None

    def is_valid_phone(self, phone):
        pattern = r'^\+?\d{11,15}$' 
        return re.match(pattern, phone) is not None

    def confirm_borrow(self, book_id):
        borrower_name = self.borrower_name_input.text().strip()
        contact_number = self.contact_input.text().strip()
        email = self.email_input.text().strip()
        gender = self.gender_input.currentText()
        classification = (
            self.other_classification_input.text().strip()
            if self.classification_input.currentText() == "Other (specify)"
            else self.classification_input.currentText()
        )

        date_borrowed = datetime.now().strftime("%Y-%m-%d %H:%M")

        if not borrower_name or not contact_number or not email:
            QMessageBox.warning(self, "Error", "All borrower details are required!")
            self.borrow_window.raise_()
            self.borrow_window.activateWindow()
            return

        if not self.is_valid_email(email):
            QMessageBox.warning(self, "Invalid Email", "Please enter a valid email address!")
            self.borrow_window.raise_()
            self.borrow_window.activateWindow()
            return

        if not self.is_valid_phone(contact_number):
            QMessageBox.warning(self, "Invalid Phone Number", "Please enter a valid phone number (digits only)!")
            self.borrow_window.raise_()
            self.borrow_window.activateWindow()
            return

        cursor.execute("SELECT AVAILABLE_COPIES FROM Library WHERE BK_ID = ?", (book_id,))
        book_data = cursor.fetchone()

        if not book_data:
            QMessageBox.warning(self, "Error", "Book not found in the database!")
            return

        available_copies = book_data[0]

        if available_copies <= 0:
            QMessageBox.warning(self, "Error", "No available copies left for this book!")
            return

        try:
            cursor.execute(
                """INSERT INTO Borrowers (BK_ID, BORROWER_NAME, CONTACT_NUMBER, EMAIL, GENDER, CLASSIFICATION, DATE_BORROWED) 
                VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (book_id, borrower_name, contact_number, email, gender, classification, date_borrowed)
            )

            available_copies -= 1
            new_status = "Available" if available_copies > 0 else "Fully Issued"

            cursor.execute("UPDATE Library SET AVAILABLE_COPIES = ?, BK_STATUS = ? WHERE BK_ID = ?", 
                (available_copies, new_status, book_id))

            connector.commit()
            QMessageBox.information(self, "Success", "Book borrowed successfully!")
            self.borrow_window.close()
            self.load_records()
            
            if hasattr(self.main_window, 'borrower_reports_widget'):
                self.main_window.borrower_reports_widget.load_borrower_reports()
                
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Error", f"Failed to borrow book: {e}")


    def remove_record(self):
        selected_row = self.table.currentRow()

        if selected_row == -1:
            QMessageBox.warning(self, "Error", "Please select a record to delete!")
            return

        book_id = self.table.item(selected_row, 1).text()  

        confirm = QMessageBox.question(
            self, "Confirm Deletion",
            f"Are you sure you want to delete this book (ID: {book_id}) and its borrower details?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if confirm == QMessageBox.Yes:
            try:
                cursor.execute("DELETE FROM Borrowers WHERE BK_ID = ?", (book_id,))  

                cursor.execute("DELETE FROM Library WHERE BK_ID = ?", (book_id,))  

                connector.commit()
                self.table.removeRow(selected_row)
                QMessageBox.information(self, "Success", "Book and borrower details deleted successfully!")

                if hasattr(self, 'borrower_window') and self.borrower_window.isVisible():
                    self.borrower_window.load_borrowers()

                if hasattr(self, 'update_dashboard'):
                    self.update_dashboard()

            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to delete record: {str(e)}")


    def remove_all_records(self):
        confirm = QMessageBox.question(self, "Confirm Deletion",
                                    "Are you sure you want to delete ALL records? This action cannot be undone.",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if confirm == QMessageBox.Yes:
            try:
                cursor.execute("DELETE FROM Library")

                cursor.execute("DELETE FROM Borrowers")

                connector.commit()
                self.table.setRowCount(0)
                QMessageBox.information(self, "Success", "All records deleted successfully!")

                if hasattr(self, 'borrower_window') and self.borrower_window.isVisible():
                    self.borrower_window.load_borrowers()

                if hasattr(self, 'update_dashboard'):
                    self.update_dashboard()

            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to delete records: {str(e)}")


    def clear_fields(self):
        self.search_input.clear()
        self.load_records()

class BorrowerReportsWidget(QWidget):
    def __init__(self, main_window, cursor, connector):
        super().__init__(main_window)
        self.main_window = main_window
        self.cursor = cursor
        self.connector = connector
        self.initUI()
        self.load_borrower_reports()
        
        self.background_pixmap = QPixmap(resource_path("images/librarybg.jpg"))

    def paintEvent(self, event):
        painter = QPainter(self)
        if not self.background_pixmap.isNull():
            scaled_pixmap = self.background_pixmap.scaled(self.size(), aspectRatioMode=1)  
            painter.drawPixmap(self.rect(), scaled_pixmap)
        super().paintEvent(event)

    def initUI(self):
        self.setStyleSheet("""
            QWidget {
                background-color: white;
                font-family: Helvetica;
            }
            QTableWidget {
                border: 2px solid #219150;
                border-radius: 5px;
                background-color: white;
                gridline-color: #E0E0E0;
                selection-background-color: #A2D9B1;
            }
            QHeaderView::section {
                background-color: #219150;
                color: white;
                font-weight: bold;
                padding: 8px;
                text-align: center;
                border: none;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:selected {
                color: #000000;
            }
            QPushButton {
                background-color: #219150;
                color: white;
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
                min-width: 160px;
            }
            QPushButton:hover {
                background-color: #1A6F33;
            }
            QPushButton#deleteButton {
                background-color: #219150;
            }
            QPushButton#deleteButton:hover {
                background-color: #C0392B;
            }
            QComboBox {
                border: 2px solid #219150;
                border-radius: 5px;        
                padding: 5px;            
                background: white;      
                font-size: 14px;
                color: #2C3E50;
            }
            QComboBox::drop-down {
                border: none; 
                width: 30px;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #219150;
                selection-background-color: #A2D9B1;
                background: white;
                color: #2C3E50;
            }
            QCheckBox {
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QCheckBox::indicator:unchecked {
                border: 2px solid #219150;
                background-color: white;
                border-radius: 3px;
            }
            QCheckBox::indicator:checked {
                border: 2px solid #219150;
                background-color: #219150;
                border-radius: 3px;
            }
        """)


        main_layout = QVBoxLayout(self)

        header_label = QLabel("Borrower Reports")
        header_label.setFont(QFont("Helvetica", 18, QFont.Bold))
        header_label.setStyleSheet("color: white; background-color: rgba(0, 0, 0, 80); padding: 10px; border-radius: 3px;")
        main_layout.addWidget(header_label)

        filter_layout = QHBoxLayout()

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["All", "This Day", "This Week", "This Month"])
        self.sort_combo.setFixedWidth(170)
        self.sort_combo.currentIndexChanged.connect(self.load_borrower_reports)

        self.status_combo = QComboBox()
        self.status_combo.addItems(["All", "Returned", "Not Returned"])
        self.status_combo.setFixedWidth(170)
        self.status_combo.currentIndexChanged.connect(self.load_borrower_reports)

        sort_label = QLabel("Filter by:")
        sort_label.setFont(QFont("Helvetica", 11, QFont.Bold))
        sort_label.setStyleSheet("color: #219150; margin-right: 6px;")

        status_label = QLabel("Return Status:")
        status_label.setFont(QFont("Helvetica", 11, QFont.Bold))
        status_label.setStyleSheet("color: #219150; margin-left: 15px; margin-right: 6px;")

        filter_layout.addWidget(sort_label)
        filter_layout.addWidget(self.sort_combo)
        filter_layout.addWidget(status_label)
        filter_layout.addWidget(self.status_combo)

        self.selection_layout = QHBoxLayout()
        
        self.select_all_button = QPushButton("✅ Select All")
        self.select_all_button.setFixedWidth(120)
        self.select_all_button.clicked.connect(self.select_all_rows)
        
        self.deselect_all_button = QPushButton("❌ Deselect All")
        self.deselect_all_button.setFixedWidth(120)
        self.deselect_all_button.clicked.connect(self.deselect_all_rows)
        
        self.selection_layout.addWidget(self.select_all_button)
        self.selection_layout.addWidget(self.deselect_all_button)
        self.selection_layout.addStretch()
        
        self.select_all_button.hide()
        self.deselect_all_button.hide()
        
        main_layout.addLayout(filter_layout)
        main_layout.addLayout(self.selection_layout)

        self.reports_table = QTableWidget()
        self.reports_table.setColumnCount(11)
        self.reports_table.setHorizontalHeaderLabels([
            "", "Borrower ID", "Book ID", "Book Title", "Borrower Name",
            "Contact", "Email", "Gender", "Classification", "Date Borrowed", "Date Returned"
        ])
        self.reports_table.setAlternatingRowColors(True)

        self.reports_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents) 
        self.reports_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents) 
        self.reports_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch) 
        self.reports_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents) 
        self.reports_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)  
        self.reports_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)  
        self.reports_table.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeToContents)  
        self.reports_table.horizontalHeader().setSectionResizeMode(10, QHeaderView.ResizeToContents)  

        self.item_changed_connections = []
        self.reports_table.itemChanged.connect(self.check_selection_status)

        main_layout.addWidget(self.reports_table)

        self.return_book_button = QPushButton("↩️ Return Book")
        self.return_book_button.clicked.connect(self.return_book_from_report)

        self.export_button = QPushButton("📤 Export to Excel")
        self.export_button.clicked.connect(self.export_to_excel)

        self.edit_button = QPushButton("✏️ Edit Borrower Details")
        self.edit_button.clicked.connect(self.edit_borrower_details)

        self.delete_button = QPushButton("🗑️ Delete Report")
        self.delete_button.setObjectName("deleteButton")
        self.delete_button.clicked.connect(self.delete_borrower_report)

        self.clear_button = QPushButton("🔄 Clear Fields")
        self.clear_button.clicked.connect(self.clear_fields)

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.return_book_button)
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addStretch()

        main_layout.addLayout(button_layout)
        
        self.has_selected_items = False
        
    def load_borrower_reports(self):
        if hasattr(self, 'item_changed_connected') and self.item_changed_connected:
            self.reports_table.itemChanged.disconnect(self.check_selection_status)
            self.item_changed_connected = False
            
        self.reports_table.setRowCount(0)

        sort_option = self.sort_combo.currentText().strip()
        status_option = self.status_combo.currentText().strip()

        query = """
            SELECT b.BORROWER_ID, b.BK_ID, l.BK_NAME, b.BORROWER_NAME, b.CONTACT_NUMBER, 
                b.EMAIL, b.GENDER, b.CLASSIFICATION, b.DATE_BORROWED, b.DATE_RETURNED
            FROM Borrowers b
            JOIN Library l ON b.BK_ID = l.BK_ID
            WHERE 1=1
        """ 

        if sort_option == "This Day":
            query += " AND substr(b.DATE_BORROWED, 1, 10) = DATE('now')"
        elif sort_option == "This Week":
            query += " AND substr(b.DATE_BORROWED, 1, 10) >= DATE('now', '-6 days')"
        elif sort_option == "This Month":
            query += " AND substr(b.DATE_BORROWED, 1, 7) = strftime('%Y-%m', 'now')"

        if status_option == "Returned":
            query += " AND b.DATE_RETURNED IS NOT NULL"
        elif status_option == "Not Returned":
            query += " AND b.DATE_RETURNED IS NULL"

        query += " ORDER BY b.BORROWER_ID ASC"

        try:
            self.cursor.execute(query)
            borrowers = self.cursor.fetchall()

            for row_num, borrower in enumerate(borrowers):
                self.reports_table.insertRow(row_num)

                checkbox_item = QTableWidgetItem()
                checkbox_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                checkbox_item.setCheckState(Qt.Unchecked)
                self.reports_table.setItem(row_num, 0, checkbox_item)

                date_returned = borrower[9] 
                is_returned = date_returned is not None and date_returned.strip() != ""

                for col_num, data in enumerate(borrower):
                    item = QTableWidgetItem(str(data))

                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

                    item.setToolTip(str(data))

                    if is_returned:
                        item.setForeground(QColor(169, 169, 169)) 

                    self.reports_table.setItem(row_num, col_num + 1, item)

            self.reports_table.itemChanged.connect(self.check_selection_status)
            self.item_changed_connected = True

            self.select_all_button.hide()
            self.deselect_all_button.hide()
            self.export_button.setText("📤 Export to Excel")
            self.has_selected_items = False

        except sqlite3.Error as e:
            QMessageBox.critical(self.main_window, "Error", f"Failed to load borrower reports: {str(e)}")

            
    def check_selection_status(self, item):
        if item.column() != 0:
            return
            
        selected_count = 0
        for row in range(self.reports_table.rowCount()):
            if self.reports_table.item(row, 0).checkState() == Qt.Checked:
                selected_count += 1
                
        if selected_count > 0:
            self.has_selected_items = True
            self.export_button.setText(f"📤 Export Selected ({selected_count})")
            self.delete_button.setText(f"🗑️ Delete Selected ({selected_count})")
            self.select_all_button.show()
            self.deselect_all_button.show()
        else:
            self.has_selected_items = False
            self.export_button.setText("📤 Export to Excel")
            self.delete_button.setText("🗑️ Delete")
            self.select_all_button.hide()
            self.deselect_all_button.hide()
            
            
    def select_all_rows(self):
        for row in range(self.reports_table.rowCount()):
            self.reports_table.item(row, 0).setCheckState(Qt.Checked)
            
    def deselect_all_rows(self):
        for row in range(self.reports_table.rowCount()):
            self.reports_table.item(row, 0).setCheckState(Qt.Unchecked)
            
    def get_selected_rows(self):
        selected_rows = []
        for row in range(self.reports_table.rowCount()):
            if self.reports_table.item(row, 0).checkState() == Qt.Checked:
                selected_rows.append(row)
        return selected_rows
    
    def export_to_excel(self):
        selected_rows = self.get_selected_rows()
        
        if not selected_rows and self.reports_table.rowCount() > 0:
            self.export_data("All")
        elif selected_rows:
            self.export_selected_rows(selected_rows)
        else:
            QMessageBox.warning(self, "Export", "No data to export")

    def export_selected_rows(self, selected_rows):
        try:
            data = []
            
            for row_index in selected_rows:
                row_data = []
                for col in range(1, self.reports_table.columnCount()):
                    item = self.reports_table.item(row_index, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
            
            columns = ["Borrower ID", "Book ID", "Book Title", "Borrower Name", 
                    "Contact", "Email", "Gender", "Classification", "Date Borrowed", "Date Returned"]
            df = pd.DataFrame(data, columns=columns)
            
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
            
            if file_path:
                df.to_excel(file_path, index=False)
                QMessageBox.information(self.main_window, "Success", f"{len(selected_rows)} records exported successfully to {file_path}")
        
        except Exception as e:
            QMessageBox.critical(self.main_window, "Error", f"Failed to export selected data: {e}")

    def export_data(self, option):
        try:
            if option == "All":
                query = """
                    SELECT b.BORROWER_ID, b.BK_ID, l.BK_NAME, b.BORROWER_NAME, b.CONTACT_NUMBER, 
                        b.EMAIL, b.GENDER, b.CLASSIFICATION, b.DATE_BORROWED, b.DATE_RETURNED
                    FROM Borrowers b
                    JOIN Library l ON b.BK_ID = l.BK_ID
                    ORDER BY b.BORROWER_ID ASC
                """

            self.cursor.execute(query)
            data = self.cursor.fetchall()

            columns = ["Borrower ID", "Book ID", "Book Title", "Borrower Name", "Contact", "Email", "Gender", "Classification", "Date Borrowed", "Date Returned"]
            df = pd.DataFrame(data, columns=columns)

            file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")

            if file_path:
                df.to_excel(file_path, index=False)
                QMessageBox.information(self.main_window, "Success", f"Data exported successfully to {file_path}")
        except Exception as e:
            QMessageBox.critical(self.main_window, "Error", f"Failed to export data: {e}")

    def return_book_from_report(self):
        selected_rows = self.get_selected_rows()
        
        if not selected_rows:
            selected_row = self.reports_table.currentRow()
            if selected_row == -1:
                QMessageBox.warning(self, "Error", "Please select a borrower record to return!")
                return
            
            selected_rows = [selected_row]
        
        if len(selected_rows) > 1:
            QMessageBox.warning(self, "Error", "Please select only one borrower record to return.")
            return
        
        selected_row = selected_rows[0]

        borrower_id = self.reports_table.item(selected_row, 1).text()
        book_id = self.reports_table.item(selected_row, 2).text()
        borrower_name = self.reports_table.item(selected_row, 4).text()

        if self.reports_table.item(selected_row, 10).text().strip() != "None" and self.reports_table.item(selected_row, 10).text().strip() != "":
            QMessageBox.information(self, "Already Returned", "This book has already been returned.")
            return

        self.return_window = QWidget()
        self.return_window.setWindowTitle("Return Book")
        
        screen = QDesktopWidget().screenGeometry()
        window_width = 400
        window_height = 150
        x = (screen.width() - window_width) // 2
        y = (screen.height() - window_height) // 2
        self.return_window.setGeometry(x, y, window_width, window_height)

        self.return_window.setStyleSheet("""
            QWidget {
                background-color: #F9F9F9;
                font-family: Helvetica;
            }
            QLabel {
                font-size: 14px;
                color: #2C3E50;
            }
            QPushButton {
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
                color: white;
            }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        header_label = QLabel(f"Return book for {borrower_name}?")
        header_label.setFont(QFont("Helvetica", 14, QFont.Bold))
        header_label.setStyleSheet("color: #219150; margin-bottom: 15px;")
        layout.addWidget(header_label)

        button_layout = QHBoxLayout()
        return_button = QPushButton("↩️ Confirm Return")
        return_button.setStyleSheet("background-color: #219150;")
        return_button.clicked.connect(lambda: self.process_return_from_report(borrower_id, book_id))
        button_layout.addWidget(return_button)

        cancel_button = QPushButton("❌ Cancel")
        cancel_button.setStyleSheet("background-color: #E74C3C;")
        button_layout.addWidget(cancel_button)
        cancel_button.clicked.connect(self.return_window.close)

        layout.addLayout(button_layout)
        self.return_window.setLayout(layout)
        self.return_window.show()

    def process_return_from_report(self, borrower_id, book_id):
        try:
            self.cursor.execute("SELECT BK_ID, TOTAL_COPIES, AVAILABLE_COPIES FROM Library WHERE BK_ID = ?", (book_id,))
            book_data = self.cursor.fetchone()

            print(f"Book ID being queried: {book_id}")
            print(f"Book data retrieved: {book_data}")

            if not book_data:
                QMessageBox.warning(self, "Error", f"Book ID {book_id} not found in the database!", QMessageBox.Ok)
                return

            total_copies, available_copies = book_data[1], book_data[2]
            
            if available_copies < total_copies:
                available_copies += 1
            else:
                QMessageBox.warning(self.main_window, "Error", "All copies of this book are already available!", QMessageBox.Ok)
                return

            date_returned = datetime.now().strftime("%Y-%m-%d %H:%M")
            self.cursor.execute("UPDATE Borrowers SET DATE_RETURNED = ? WHERE BORROWER_ID = ?", (date_returned, borrower_id))
            self.cursor.execute("UPDATE Library SET AVAILABLE_COPIES = ?, BK_STATUS = ? WHERE BK_ID = ?", 
                (available_copies, "Available" if available_copies > 0 else "Fully Issued", book_id))
            
            self.connector.commit()

            QMessageBox.information(self.main_window, "Success", "Book returned successfully!", QMessageBox.Ok)
            self.return_window.close()
            self.load_borrower_reports()

            if hasattr(self.main_window, 'inventory_widget'):
                self.main_window.inventory_widget.load_records()

        except sqlite3.Error as e:
            QMessageBox.critical(self, "Error", f"Failed to return book: {e}", QMessageBox.Ok)


    def edit_borrower_details(self):
        selected_rows = self.get_selected_rows()
        
        if not selected_rows:
            selected_row = self.reports_table.currentRow()
            if selected_row == -1:
                QMessageBox.warning(self.main_window, "Error", "Please select a borrower to edit!")
                return
            
            selected_rows = [selected_row]
        
        if len(selected_rows) > 1:
            QMessageBox.warning(self.main_window, "Error", "Please select only one borrower to edit.")
            return
        
        selected_row = selected_rows[0]
        
        date_returned = self.reports_table.item(selected_row, 10).text()

        if date_returned.lower() == "none" or not date_returned.strip():
            borrower_id = self.reports_table.item(selected_row, 1).text()
            borrower_name = self.reports_table.item(selected_row, 4).text()
            contact = self.reports_table.item(selected_row, 5).text()
            email = self.reports_table.item(selected_row, 6).text()
            gender = self.reports_table.item(selected_row, 7).text()
            classification_value = self.reports_table.item(selected_row, 8).text()

            self.edit_window = QWidget()
            self.edit_window.setWindowTitle("Edit Borrower Details")
            self.edit_window.setGeometry(400, 200, 420, 400)

            self.edit_window.setStyleSheet("""
                QWidget {
                    background-color: white;
                    font-family: Helvetica;
                }
                QLineEdit, QComboBox {
                    border: 2px solid #219150;
                    border-radius: 5px;
                    padding: 8px;
                    font-size: 14px;
                    background-color: white;
                }
                QPushButton {
                    border-radius: 5px;
                    padding: 10px;
                    font-size: 14px;
                    color: white;
                }
                QLabel {
                    font-size: 14px;
                    color: #2C3E50;
                }
            """)

            layout = QVBoxLayout()
            form_layout = QFormLayout()

            self.edit_name = QLineEdit(borrower_name)
            self.edit_contact = QLineEdit(contact)
            self.edit_email = QLineEdit(email)

            self.edit_gender = QComboBox()
            self.edit_gender.addItems(["Male", "Female"])
            self.edit_gender.setCurrentText(gender)

            self.edit_classification = QComboBox()
            classification_options = ["Student", "Faculty", "REPS", "Other"]
            self.edit_classification.addItems(classification_options)

            self.other_classification_container = QWidget()
            other_layout = QHBoxLayout(self.other_classification_container)
            other_layout.setContentsMargins(0, 0, 0, 0)

            other_label = QLabel("Other (Specify):")
            self.other_classification_input = QLineEdit()
            self.other_classification_input.setPlaceholderText("Please specify...")
            self.other_classification_input.setMaximumWidth(280)

            other_layout.addWidget(other_label)
            other_layout.addWidget(self.other_classification_input)

            classification_value = self.reports_table.item(selected_row, 8).text()

            if classification_value in classification_options:
                self.edit_classification.setCurrentText(classification_value)
                self.other_classification_container.hide()
            else:
                self.edit_classification.setCurrentText("Other")
                self.other_classification_input.setText(classification_value)
                self.other_classification_container.show()

            self.edit_classification.currentTextChanged.connect(
                lambda value: self.other_classification_container.setVisible(value == "Other")
            )

            form_layout.addRow("Borrower Name:", self.edit_name)
            form_layout.addRow("Contact Number:", self.edit_contact)
            form_layout.addRow("Email Address:", self.edit_email)
            form_layout.addRow("Gender:", self.edit_gender)
            form_layout.addRow("Classification:", self.edit_classification)
            form_layout.addRow(self.other_classification_container)

            layout.addLayout(form_layout)

            button_layout = QHBoxLayout()

            self.save_button = QPushButton("💾 Save")
            self.save_button.setStyleSheet("background-color: #219150; color: white; padding: 10px; min-width: 120px;")
            self.save_button.clicked.connect(lambda: self.save_borrower_details(borrower_id))
            button_layout.addWidget(self.save_button)

            self.cancel_button = QPushButton("❌ Cancel")
            self.cancel_button.setStyleSheet("background-color: #E74C3C; color: white; padding: 10px; min-width: 120px;")
            self.cancel_button.clicked.connect(self.edit_window.close)
            button_layout.addWidget(self.cancel_button)

            layout.addLayout(button_layout)
            self.edit_window.setLayout(layout)
            self.edit_window.show()

        else:  
            QMessageBox.information(
                self.main_window,
                "Cannot Edit",
                "This borrower record is linked to a returned book and cannot be edited."
            )
            return

    def save_borrower_details(self, borrower_id):
        new_name = self.edit_name.text().strip()
        new_contact = self.edit_contact.text().strip()
        new_email = self.edit_email.text().strip()
        new_gender = self.edit_gender.currentText()
        if self.edit_classification.currentText() == "Other":
            new_classification = self.other_classification_input.text().strip()
        else:
            new_classification = self.edit_classification.currentText()


        if not new_name:
            QMessageBox.warning(self.main_window, "Invalid Name", "Borrower name cannot be empty!")
            self.edit_window.raise_()
            self.edit_window.activateWindow()
            return

        if not self.is_valid_phone(new_contact):
            QMessageBox.warning(self.main_window, "Invalid Phone Number", "Please enter a valid phone number (10-15 digits)!")
            self.edit_window.raise_()
            self.edit_window.activateWindow()
            return

        if not self.is_valid_email(new_email):
            QMessageBox.warning(self.main_window, "Invalid Email", "Please enter a valid email address!")
            self.edit_window.raise_()
            self.edit_window.activateWindow()
            return

        try:
            self.cursor.execute("""
                UPDATE Borrowers 
                SET BORROWER_NAME=?, CONTACT_NUMBER=?, EMAIL=?, GENDER=?, CLASSIFICATION=? 
                WHERE BORROWER_ID=?
            """, (new_name, new_contact, new_email, new_gender, new_classification, borrower_id))

            self.connector.commit()
            QMessageBox.information(self.main_window, "Success", "Borrower details updated successfully!")
            self.edit_window.close()
            self.load_borrower_reports()

        except sqlite3.Error as e:
            QMessageBox.critical(self, "Error", f"Failed to update borrower details: {e}")

    def delete_borrower_report(self):
        selected_rows = self.get_selected_rows()
        
        if not selected_rows:
            selected_row = self.reports_table.currentRow()
            if selected_row == -1:
                QMessageBox.warning(self.main_window, "No Selection", "Please select a borrower report to delete.", QMessageBox.Ok)
                return
            
            selected_rows = [selected_row]
        
        selected_borrower_ids = []
        for row in selected_rows:
            borrower_id = self.reports_table.item(row, 1).text()
            selected_borrower_ids.append(borrower_id)
        
        if len(selected_borrower_ids) == 1:
            message = f"Are you sure you want to delete Borrower ID {selected_borrower_ids[0]}?"
        else:
            message = f"Are you sure you want to delete {len(selected_borrower_ids)} selected borrower records?"
        
        confirmation = QMessageBox.question(
            self.main_window,
            "Delete Confirmation",
            message,
            QMessageBox.Yes | QMessageBox.No
        )
        
        if confirmation == QMessageBox.No:
            return
        
        try:
            for borrower_id in selected_borrower_ids:
                self.cursor.execute("DELETE FROM Borrowers WHERE BORROWER_ID = ?", (borrower_id,))
            
            self.connector.commit()
            self.reorder_borrower_ids()
            self.load_borrower_reports()
            
            if len(selected_borrower_ids) == 1:
                QMessageBox.information(self.main_window, "Success", "Borrower report deleted successfully.", QMessageBox.Ok)
            else:
                QMessageBox.information(self.main_window, "Success", f"{len(selected_borrower_ids)} borrower reports deleted successfully.", QMessageBox.Ok)
        
        except sqlite3.Error as e:
            QMessageBox.critical(self.main_window, "Error", f"Failed to delete borrower report(s): {e}", QMessageBox.Ok)

    def reorder_borrower_ids(self):
        try:
            self.cursor.execute("SELECT BORROWER_ID FROM Borrowers ORDER BY BORROWER_ID ASC")
            borrowers = self.cursor.fetchall()

            new_id = 1
            for borrower in borrowers:
                old_id = borrower[0]
                self.cursor.execute("UPDATE Borrowers SET BORROWER_ID = ? WHERE BORROWER_ID = ?", (new_id, old_id))
                new_id += 1

            self.connector.commit()

            self.cursor.execute("DELETE FROM sqlite_sequence WHERE name='Borrowers'")
            self.connector.commit()

            print("Borrower IDs reordered successfully.")

        except sqlite3.Error as e:
            print(f"Error reordering borrower IDs: {e}")

    def is_valid_phone(self, phone):
        digits_only = ''.join(filter(str.isdigit, phone))
        return 10 <= len(digits_only) <= 15

    def is_valid_email(self, email):
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None

    def clear_fields(self):
        self.sort_combo.setCurrentIndex(0)
        self.status_combo.setCurrentIndex(0)
        
        self.deselect_all_rows()
        
        self.load_borrower_reports()
        
        self.has_selected_items = False
        self.export_button.setText("📤 Export to Excel")
        self.delete_button.setText("🗑️ Delete Report")
        self.select_all_button.hide()
        self.deselect_all_button.hide()

class LibraryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.connector = sqlite3.connect("library.db")
        self.cursor = self.connector.cursor()
        self.initUI()
        

    def initUI(self):
        self.setWindowTitle("Library Management System")
        self.setGeometry(100, 100, 1200, 700)
        self.setMinimumSize(800, 600)

        self.setStyleSheet("""
            QMainWindow {
                background-color: #F5F5F5;
            }
        """)
 
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
     
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        self.create_sidebar()
        
        self.content_stack = QStackedWidget()
        self.main_layout.addWidget(self.content_stack)
        
        self.dashboard_widget = DashboardWidget(self.cursor, self.connector)
        self.content_stack.addWidget(self.dashboard_widget)
        
        self.inventory_widget = InventoryWidget(self, self.cursor, self.connector)
        self.content_stack.addWidget(self.inventory_widget)

        # Add this after self.inventory_widget:
        self.borrower_reports_widget = BorrowerReportsWidget(self, self.cursor, self.connector)
        self.content_stack.addWidget(self.borrower_reports_widget)
        
        self.content_stack.setCurrentIndex(0)
        self.dashboard_button.setChecked(True)
        
    def create_sidebar(self):
        self.sidebar = QWidget()
        self.sidebar.setFixedWidth(230)
        self.sidebar.setStyleSheet("""
            QWidget {
                background-color: #2C3E50;
            }
        """)
        self.main_layout.addWidget(self.sidebar)
        
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)
        
        logo_container = QWidget()
        logo_container.setStyleSheet("background-color: #1A2530; padding: 15px;")
        logo_container.setFixedHeight(80)
        logo_layout = QHBoxLayout(logo_container)
        
        logo_label = QLabel("📚 CISC")
        logo_label.setStyleSheet("color: white; font-size: 22px; font-weight: bold;")
        logo_layout.addWidget(logo_label)
        
        sidebar_layout.addWidget(logo_container)
        
        self.dashboard_button = SidebarButton("  Dashboard", resource_path("images/dashboard.png"))
        self.dashboard_button.clicked.connect(lambda: self.change_page(0))
        sidebar_layout.addWidget(self.dashboard_button)
        
        self.inventory_button = SidebarButton("  Inventory", resource_path("images/library.png"))
        self.inventory_button.clicked.connect(lambda: self.change_page(1))
        sidebar_layout.addWidget(self.inventory_button)

        self.borrowers_button = SidebarButton("  Borrower Reports", resource_path("images/reports.png"))
        self.borrowers_button.clicked.connect(lambda: self.change_page(2))
        sidebar_layout.addWidget(self.borrowers_button)
        
        sidebar_layout.addStretch()
        
        version_label = QLabel("Version 1.0.0")
        version_label.setStyleSheet("color: #95A5A6; padding: 10px;")
        version_label.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(version_label)
        
    def change_page(self, index):
        self.dashboard_button.setChecked(False)
        self.inventory_button.setChecked(False)
        self.borrowers_button.setChecked(False)

        self.content_stack.setCurrentIndex(index)

        if index == 0:
            self.dashboard_button.setChecked(True)
            self.dashboard_widget.refresh_data()
        elif index == 1:
            self.inventory_button.setChecked(True)
            self.inventory_widget.load_records()
        elif index == 2:
            self.borrowers_button.setChecked(True)

if __name__ == "__main__":
    import sys
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path("images/cpaflogo.png")))
    window = LibraryApp()
    window.show()
    sys.exit(app.exec_())
