# Xlreport – Easy Excel Export for Python Objects
[![Downloads](https://img.shields.io/pypi/dm/xlreport.svg)](https://pypi.org/project/xlreport/) [![PyPI version](https://img.shields.io/pypi/v/xlreport.svg)](https://pypi.org/project/xlreport/)

Xlreport is a Python wrapper around [xlsxwriter](https://xlsxwriter.readthedocs.io/) that makes it super easy to dump your data into Excel files. It takes care of all the formatting, so you don't have to mess with xlsxwriter's details. Most of the time, you'll just use the `to_file` function to throw whatever data you've got into a nicely formatted spreadsheet.

While Xlreport is mainly designed for 2D data (think tables with rows and columns), it can handle 1D data too. If you pass in a single-dimension list or array, it'll just fill up one column in your Excel file - simple as that.

## Installation

```bash
pip install xlreport
```

## Quick Start

```python
# Generate some random data
from numpy.random import default_rng
arr = default_rng(42).random((100, 4))

# Save data to an Excel file and immediately open it
import xlreport as xl
header = ['col1', 'col2', 'col3', 'col4'] 
xl.to_file("test.xlsx", arr, header, title="Test numpy")
```

## Supported Data Types

Xlreport works with various Python data structures:
- **Lists and tuples** - Regular Python sequences
- **NumPy arrays** - Numerical arrays with automatic conversion
- **Pandas DataFrames** - With automatic index and column handling
- **Sets** - Converted to list format
- **1D data** - Single columns for simple lists

### 📊 **Multi-Sheet Excel Files**
Create complex reports with multiple sheets and navigation links:

```python
import xlreport as xl

# Create a multi-sheet file
exfile = xl.Exfile("report.xlsx")

# Add different datasets to different sheets
exfile.write(sales_data, title="Sales Report", worksheet_name="Sales")
exfile.write(customer_data, title="Customer Data", worksheet_name="Customers")
exfile.write(system_info, title="System Info", worksheet_name="Info", wrap=True)

# Add navigation links between sheets
exfile.add_links()

# Save and open the file
exfile.save(start=True)
```

![Sample multisheets spreadsheet](https://github.com/ts-kontakt/xlreport/blob/main/xlreport-gnumeric.gif?raw=true)



### 📏 **Intelligent Column Sizing**
Xlreport addresses common challenges with Excel's automatic column sizing. Unlike xlsxwriter's `worksheet.autofit()` method which can produce undesirable results (like excessively wide columns for long strings), Xlreport offers a smart solution.

The column width is automatically calculated based on header text length using a logarithmic formula that provides optimal readability. You can fine-tune column widths by simply adding spaces to your headers:

```python
# Control column width with spaces
headers = [
    "ID",                                    # Narrow column
    "Name          ",                        # Medium column  
    "Description - longer column          "  # Wide column
]
xl.to_file("sized.xlsx", data, headers, title="Custom Sizing")
```

### 🎨 **Professional Formatting**
- **Frozen headers** - Top row stays visible when scrolling
- **Automatic number formatting** - Decimals, integers, and currency
- **Text wrapping** - Long text content with `wrap=True` option
- **Color-coded headers** - Professional appearance with consistent styling
- **Unicode support** - Handles international characters correctly

### 🔗 **Sheet Navigation**
When creating multi-sheet files, `add_links()` automatically creates clickable navigation links at the top of each sheet, making it easy to jump between different data views.

## Advanced Usage

### Working with Pandas DataFrames

```python
import pandas as pd
import xlreport as xl

df = pd.DataFrame({
    "Product": ["Widget A", "Widget B", "Widget C"],
    "Sales": [1000, 1500, 800],
    "Profit": [200.50, 350.75, 120.25]
})

xl.to_file("dataframe.xlsx", df, title="Product Sales")
```

### Text Wrapping for Long Content

```python
# For data with long text content
long_text_data = [
    ["Topic", "Description"],
    ["AI", "Artificial Intelligence is a branch of computer science..."],
    ["ML", "Machine Learning is a subset of AI that focuses on..."]
]

xl.to_file("docs.xlsx", long_text_data, wrap=True, title="Documentation")
```

### Custom Formatting Example

```python
import xlreport as xl

# Create file with custom sheet name and formatting
exfile = xl.Exfile("custom_report.xlsx")

# Add data with specific formatting
exfile.write(
    data=financial_data,
    title="Q4 Financial Report",
    worksheet_name="Q4_Financials",
    wrap=False
)

# Add links and save
exfile.add_links()
exfile.save(start=True)  # Automatically opens the file
```

## Dependencies

- **[xlsxwriter](https://xlsxwriter.readthedocs.io/)** - The core Excel writing library

## Use Cases

Xlreport is perfect for:

- **Data inspection** - Quick visual analysis of datasets
- **Report generation** - Professional-looking Excel reports
- **Data sharing** - Easy export for stakeholders
- **Prototyping** - Rapid data visualization during development
- **Multi-sheet reports** - Complex documents with navigation

## Alternative: HTML Tables

If you don't like spreadsheets there is better solution, check out **[df2tables](https://github.com/ts-kontakt/df2tables)** - a companion library that generates interactive HTML tables from your data.


## License

This code is licensed under MIT

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

---

*Need more control over Excel formatting? Check out the full [xlsxwriter documentation](https://xlsxwriter.readthedocs.io/) for advanced features.*
