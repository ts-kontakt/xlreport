## Nice formatted excel file from iterables (preferably list)
* [General info](#general-info)
* [Technologies](#technologies)
* [Setup](#setup)

## General info
Generating nicely formatted excel file as easy as:
```python
import numpy as np
from numpy.random import default_rng
arr = default_rng(42).random((100, 4))

import xlreport as xl
header = ['col1', 'col2', 'col3', 'col4']
#this is try to open file as well
xl.save_list("test.xlsx", arr.tolist(), header, title="Test numpy")
```
	
## Technologies
Project is created with:
* Lorem version: 12.3
* Ipsum version: 2.33
* Ament library version: 999

## Screenshots




<img src="KvYmeqXZcj.gif"   style="max-width: 100%;max-height: 100%;">
<!-- If you have screenshots you'd like to share, include them here. -->

	
## Setup
To run this project, install it locally using npm:
