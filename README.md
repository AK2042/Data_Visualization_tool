# Excel to Graph Converter

This repository provides two graphical applications for loading CSV data files and generating various types of visualizations. The applications are implemented in **PyQt5** and **Tkinter**, allowing users to interactively load data, select graph types, and view customized visualizations.

## Features

- **Load CSV files** for data analysis
- **View dataset** within the application interface
- **Select X and Y columns** for plotting
- **Choose from multiple plot types**, including:
  - Line Plot
  - Bar Chart
  - Scatter Plot
  - Histogram
  - Area Plot
  - Regression Plot
  - Heatmap
  - Pie Chart
- **Toggle between Light and Dark themes** (available in PyQt version)
- **View generated graphs** within the application

## Prerequisites

Before running the application, ensure you have the following Python packages installed. These are listed in `requirements.txt`:

```plaintext
pandas
matplotlib
seaborn
PyQt5
```

To install all required packages, run:

```bash
pip install -r requirements.txt
```

## Installation and Usage

### PyQt5 Version

This version provides a comprehensive GUI with a dark mode option and column selection features.

1. Run the PyQt5 application script:

    ```bash
    python main.py
    ```

2. Once launched, you can:
   - **Load a CSV file** by clicking "Load CSV File."
   - View the **dataset in a table** and **select columns** for the X and Y axes.
   - Choose a **graph type** and click "Generate Graph" to view the visualization.
   - Switch between **light and dark themes** by clicking the "Switch to Dark Mode" button.

### Tkinter Version

The Tkinter version provides a simpler interface for loading and visualizing data, with an option to select specific columns or plot all data.

1. Run the Tkinter-based application script:

    ```bash
    python tkinter_version.py
    ```

2. In this version:
   - **Load a CSV file** by clicking "Browse."
   - Select a **graph type** and specify **dataset column(s)** for plotting.
   - Click "Generate Graph" to view the visualization within the application.

## Additional Notes

- Both applications include basic **error handling** to notify users of issues with data loading or plotting.
- Future enhancements, such as **data filtering** and **custom color selection**, could be added as needed.

## License

This project is licensed under the MIT License.
