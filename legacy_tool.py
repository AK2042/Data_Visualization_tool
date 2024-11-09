import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from tkinter import Tk, Label, Button, filedialog, Canvas, StringVar, OptionMenu
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class ExcelToGraphConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to Graph Converter")

        self.file_path = StringVar()
        self.graph_type = StringVar()

        Label(root, text="Select CSV File:").grid(row=0, column=0, pady=10, padx=10)
        Button(root, text="Browse", command=self.browse_file).grid(row=0, column=1, pady=10, padx=10, sticky='w')

        Label(root, text="Select Graph Type:").grid(row=1, column=0, pady=10, padx=10)
        graph_options = ["Click->", "Line Plot", "Bar Chart", "Scatter Plot", "Histogram", "Area Plot", "Regression Plot", "Heat Map", "Pie Chart"]
        self.graph_type.set(graph_options[0])
        OptionMenu(root, self.graph_type, *graph_options).grid(row=1, column=1, pady=10, padx=10, sticky='w')

        Label(root, text="Do You Want to Enter Specific Data?(If No, Then Enter 'All',\n Else, Enter the Column name:)").grid(row=2, column=0, pady=10, padx=10)
        self.dataset = StringVar()
        Button(root, text="Show Data", command=self.show_data).grid(row=4, column=0, pady=10, padx=10)

        Button(root, text="Generate Graph", command=self.generate_graph).grid(row=4, column=1, pady=10, padx=10, sticky='w')

        self.graph_canvas = Canvas(root, width=600, height=400)
        self.graph_canvas.grid(row=5, column=0, columnspan=2, pady=10, padx=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV", ".csv")])
        self.file_path.set(file_path)

    def generate_graph(self):
        file_path = self.file_path.get()
        graph_type = self.graph_type.get()

        if not file_path or not graph_type:
            return

        df = pd.read_csv(file_path)

        plt.clf()
        fig, ax = plt.subplots()

        if graph_type == "Line Plot":
            if self.dataset.get() == 'All':
                sns.lineplot(data=df, ax=ax)
            else:
                sns.lineplot(x=df.index, y=df[self.dataset.get()], data=df, ax=ax)
        elif graph_type == "Bar Chart":
            if self.dataset.get() == 'All':
                sns.barplot(data=df, ax=ax)
            else:
                sns.barplot(x=df.index, y=df[self.dataset.get()], data=df, ax=ax)
        elif graph_type == "Scatter Plot":
            if self.dataset.get() == 'All':
                sns.scatterplot(data=df, ax=ax)
            else:
                sns.scatterplot(x=df.index, y=df[self.dataset.get()], data=df, ax=ax)
        elif graph_type == "Histogram":
            if self.dataset.get() == 'All':
                sns.histplot(data=df, ax=ax)
            else:
                sns.histplot(x=df.index, y=df[self.dataset.get()], data=df, ax=ax)
        elif graph_type == "Area Plot":
            if self.dataset.get() == 'All':
                sns.kdeplot(data=df, ax=ax)
            else:
                sns.kdeplot(x=df.index, y=df[self.dataset.get()], data=df, ax=ax)
        elif graph_type == "Regression Plot":
            if self.dataset.get() == 'All':
                sns.regplot(data=df, ax=ax)
            else:
                sns.regplot(x=df.index, y=df[self.dataset.get()], data=df, ax=ax)
        elif graph_type == "Heat Map":
            if self.dataset.get() == 'All':
                sns.heatmap(df, cmap="coolwarm", ax=ax)
            else:
                sns.heatmap(x=df.index, y=df[self.dataset.get()], data=df,  cmap="coolwarm", ax=ax)
        elif graph_type == "Pie Chart":
            if self.dataset.get() == 'All':
                plt.pie(df, ax=ax)
            else:
                plt.pie(df[self.dataset.get()], ax=ax)

        ax.set_title(graph_type)
        ax.set_xlabel(df.columns[0])
        ax.set_ylabel(df.columns[1])

        self.show_graph(fig)

    def show_data(self):
        file_path = self.file_path.get()
        Show_Data = ["click->"]
        for i in pd.read_csv(file_path,sep=","):
            if i=='':
                break
            Show_Data.append(i)
        Show_Data.append("All")
        Label(root, text="Columns in Data").grid(row=4, column=0)
        OptionMenu(root, self.dataset, *Show_Data).grid(row=2, column=1, sticky='w')


    def show_graph(self, fig):
        canvas = FigureCanvasTkAgg(fig, master=self.root)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=5, column=0, columnspan=2, pady=10, padx=10)
        canvas.draw()

if __name__ == "__main__":
    root = Tk()
    app = ExcelToGraphConverter(root)
    root.mainloop()