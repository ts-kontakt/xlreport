# Xlreport – Streamlined Data Export to Excel


## Basic Functionality
Xlreport is a Python module created for quick data transfer to Excel spreadsheets. 
Xlreport has a single dependency: popular xlsxwriter package.
In most cases, it is enough to use a simple ```to_file``` function, which writes a python object (list, tuple, dataframe, etc.) to a nicely formatted excel file.



```python
#generate some random data
from numpy.random import default_rng
arr = default_rng(42).random((100, 4))


# Save data to an Excel file and immediately invoke the default application to open it (does not have to be ms excel)
import xlreport as xl
header = 'col1 col2 col3 col4'.split(' ') 
xl.to_file("test.xlsx", arr, header, title="Test numpy")

```
 
It is a useful tool for data inspection, as well as for simply creating reports. It is equally easy to create a file containing multiple sheets and links to them.

## Screenshots

Example of multisheet file.
If we want to create a file containing multiple sheets, we can use the Exfile class.
By default, the header cells are frozen to make it easier to browse multiple rows of data. (see: [test_xlreport.py](test_xlreport.py))




```python

import xlreport as xl

# get some diverse datasets for demonstration
data1 = get_packages()
data2 = generate_random_data(20)
data3 = [(x, y) for x, y in system_info().items()]
data3 = [(x, y) for x, y in system_info().items()]

# Construct and populate a multi-sheet Excel file
exfile = xl.Exfile("test_multisheet_file.xlsx")
exfile.write(data1, title="Current user packages")
exfile.write(data2, title="Random data")
exfile.write(data3, title="System Info", wrap=True)
exfile.add_links()
exfile.save(start=True)
```

<p align="left">
<img src="xlreport-gnumeric.gif"   width="500" style="max-width: 100%;max-height: 100%;">
<!-- If you have screenshots you'd like to share, include them here. -->
</p>


## Intelligent Column Sizing

Xlreport addresses common challenges associated with automatic column sizing in Excel. While xlsxwriter's worksheet.autofit() method can often produce undesirable results (e.g., excessively wide columns for long strings), Xlreport offers a pragmatic solution. Users can effectively control column widths by simply appending spaces to header strings, providing a straightforward method for layout optimization:

This approach mitigates the need for manual width calculations.
   

```python
header = ["Normal column    ", "Description - longer column          "]
```

## Installation
To integrate Xlreport into your Python environment, simply place the xlreport.py file within your site-packages directory. You can locate this directory using:
```python
import site
print(site.getusersitepackages())
```

## Dependencies


 ```xlsxwriter``` .



