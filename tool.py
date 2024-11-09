import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                             QWidget, QListWidget, QLabel, QComboBox,
                             QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, 
                             QFileDialog, QDialog)
from PyQt5.QtGui import QFont
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas


class GraphWindow(QDialog):
    def __init__(self, graph_type, x_columns, y_columns, df, dark_mode):
        super().__init__()
        self.setWindowTitle("Generated Graph")
        self.setGeometry(100, 100, 1600, 1400)

        self.layout = QVBoxLayout()
        self.figure = plt.figure(figsize=(10, 6))
        self.canvas = FigureCanvas(self.figure)
        self.layout.addWidget(self.canvas)
        self.setLayout(self.layout)

        self.generate_graph(graph_type, x_columns, y_columns, df, dark_mode)

    def generate_graph(self, graph_type, x_columns, y_columns, df, dark_mode):
        self.canvas.figure.clear()
        ax = self.canvas.figure.add_subplot(111)

        if dark_mode:
            ax.set_facecolor('#2E2E2E')
            self.canvas.setStyleSheet("background-color: #2E2E2E;")
        else:
            ax.set_facecolor('white')
            self.canvas.setStyleSheet("background-color: white;")

        try:
            if graph_type == "Line Plot":
                for y_column in y_columns:
                    sns.lineplot(data=df, x=x_columns[0], y=y_column, ax=ax, label=y_column)
            elif graph_type == "Bar Chart":
                for y_column in y_columns:
                    sns.barplot(data=df, x=x_columns[0], y=y_column, ax=ax, label=y_column)
            elif graph_type == "Scatter Plot":
                for y_column in y_columns:
                    for x_column in x_columns:
                        sns.scatterplot(data=df, x=x_column, y=y_column, ax=ax, label=f"{y_column} vs {x_column}")
            elif graph_type == "Histogram":
                for x_column in x_columns:
                    sns.histplot(data=df, x=x_column, ax=ax, label=x_column)
            elif graph_type == "Area Plot":
                for x_column in x_columns:
                    sns.kdeplot(data=df, x=x_column, ax=ax, label=x_column)
            elif graph_type == "Regression Plot":
                for y_column in y_columns:
                    for x_column in x_columns:
                        sns.regplot(data=df, x=x_column, y=y_column, ax=ax, label=f"{y_column} vs {x_column}")
            elif graph_type == "Pie Chart":
                if len(y_columns) == 1:
                    df_counts = df[y_columns[0]].value_counts()
                    ax.pie(df_counts, labels=df_counts.index, autopct='%1.1f%%', startangle=90)

            ax.set_title(f"{graph_type} of {' and '.join(y_columns)} vs {' and '.join(x_columns)}")
            ax.set_xlabel(' & '.join(x_columns))
            ax.set_ylabel(' & '.join(y_columns))
            ax.legend()
            self.canvas.draw()

        except Exception as e:
            QMessageBox.warning(self, "Plotting Error", f"An error occurred while plotting: {e}")


class ExcelToGraphConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to Graph Converter")
        self.setGeometry(100, 100, 1000, 800)

        self.layout = QVBoxLayout()
        font = QFont()
        font.setBold(True)

        self.theme_toggle_button = QPushButton("Switch to Dark Mode")
        self.theme_toggle_button.setFont(font)
        self.theme_toggle_button.clicked.connect(self.toggle_theme)
        self.layout.addWidget(self.theme_toggle_button)

        self.load_button = QPushButton("Load CSV File")
        self.load_button.setFont(font)
        self.load_button.clicked.connect(self.load_data)
        self.layout.addWidget(self.load_button)

        self.table_label = QLabel("Your Dataset")
        self.table_label.setFont(font)
        self.layout.addWidget(self.table_label)
        self.data_table = QTableWidget()
        self.layout.addWidget(self.data_table)

        self.x_axis_label = QLabel("Select X Axis Columns:")
        self.x_axis_label.setFont(font)
        self.layout.addWidget(self.x_axis_label)
        self.x_axis_list = QListWidget()
        self.x_axis_list.setSelectionMode(QListWidget.MultiSelection)
        self.x_axis_list.itemSelectionChanged.connect(self.update_graph_options)
        self.layout.addWidget(self.x_axis_list)

        self.y_axis_label = QLabel("Select Y Axis Columns:")
        self.y_axis_label.setFont(font)
        self.layout.addWidget(self.y_axis_label)
        self.y_axis_list = QListWidget()
        self.y_axis_list.setSelectionMode(QListWidget.MultiSelection)
        self.y_axis_list.itemSelectionChanged.connect(self.update_graph_options)
        self.layout.addWidget(self.y_axis_list)

        self.graph_type_label = QLabel("Select Graph Type (From available Graphs given below):")
        self.graph_type_label.setFont(font)
        self.layout.addWidget(self.graph_type_label)
        self.graph_options = QComboBox()
        self.layout.addWidget(self.graph_options)

        self.generate_button = QPushButton("Generate Graph")
        self.generate_button.setFont(font)
        self.generate_button.clicked.connect(self.show_graph_window)
        self.layout.addWidget(self.generate_button)

        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

        self.df = None
        self.dark_mode = False
        self.apply_initial_theme()

    def apply_initial_theme(self):
        if self.dark_mode:
            self.set_dark_mode()
        else:
            self.set_light_mode()

    def set_dark_mode(self):
        self.setStyleSheet("background-color: #2E2E2E; color: white;")
        self.data_table.setStyleSheet("background-color: #444; color: white;")
        self.data_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #555; color: white; }")
        self.graph_options.setStyleSheet("background-color: #333; color: white;")
        self.theme_toggle_button.setText("Switch to Light Mode")

    def set_light_mode(self):
        self.setStyleSheet("")
        self.data_table.setStyleSheet("")
        self.data_table.horizontalHeader().setStyleSheet("")
        self.graph_options.setStyleSheet("")
        self.theme_toggle_button.setText("Switch to Dark Mode")

    def load_data(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv)")
        if file_path:
            self.df = pd.read_csv(file_path)
            self.display_data()

    def display_data(self):
        if self.df is not None:
            self.data_table.setRowCount(len(self.df))
            self.data_table.setColumnCount(len(self.df.columns))
            self.data_table.setHorizontalHeaderLabels(self.df.columns)

            for i in range(len(self.df)):
                for j in range(len(self.df.columns)):
                    item = QTableWidgetItem(str(self.df.iat[i, j]))
                    self.data_table.setItem(i, j, item)

            self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

            self.x_axis_list.clear()
            self.y_axis_list.clear()
            self.x_axis_list.addItems(self.df.columns)
            self.y_axis_list.addItems(self.df.columns)

            self.graph_options.clear()

    def update_graph_options(self):
        x_columns = [item.text() for item in self.x_axis_list.selectedItems()]
        y_columns = [item.text() for item in self.y_axis_list.selectedItems()]

        available_graphs = set()

        if x_columns and y_columns:
            for x_col in x_columns:
                if pd.api.types.is_numeric_dtype(self.df[x_col]):
                    available_graphs.add("Line Plot")
                    available_graphs.add("Bar Chart")
                    available_graphs.add("Scatter Plot")
                    available_graphs.add("Histogram")
                    available_graphs.add("Area Plot")
                    available_graphs.add("Regression Plot")
                elif pd.api.types.is_categorical_dtype(self.df[x_col]) or pd.api.types.is_object_dtype(self.df[x_col]):
                    available_graphs.add("Bar Chart")
                    available_graphs.add("Pie Chart")

            for y_col in y_columns:
                if pd.api.types.is_numeric_dtype(self.df[y_col]):
                    available_graphs.add("Line Plot")
                    available_graphs.add("Bar Chart")
                    available_graphs.add("Scatter Plot")
                    available_graphs.add("Histogram")
                    available_graphs.add("Area Plot")
                    available_graphs.add("Regression Plot")
                elif pd.api.types.is_categorical_dtype(self.df[y_col]) or pd.api.types.is_object_dtype(self.df[y_col]):
                    available_graphs.add("Bar Chart")
                    available_graphs.add("Pie Chart")

        self.graph_options.clear()
        self.graph_options.addItems(available_graphs)

    def show_graph_window(self):
        graph_type = self.graph_options.currentText()
        x_columns = [item.text() for item in self.x_axis_list.selectedItems()]
        y_columns = [item.text() for item in self.y_axis_list.selectedItems()]

        if not graph_type:
            QMessageBox.warning(self, "Selection Error", "Please select a graph type.")
            return

        if not x_columns or not y_columns:
            QMessageBox.warning(self, "Selection Error", "Please select at least one X and one Y column.")
            return

        graph_window = GraphWindow(graph_type, x_columns, y_columns, self.df, self.dark_mode)
        graph_window.exec_()

    def toggle_theme(self):
        if self.dark_mode:
            self.set_light_mode()
        else:
            self.set_dark_mode()
        self.dark_mode = not self.dark_mode


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    mainWin = ExcelToGraphConverter()
    mainWin.show()
    sys.exit(app.exec_())
