# SpreadPandas: data frame to spreadsheet mapping
[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/release/python-370/)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![GitHub license](https://badgen.net/github/license/FilippoPisello/EasyPred)](https://github.com/FilippoPisello/SpreadPandas/blob/main/LICENSE)
[![PyPI version fury.io](https://badge.fury.io/py/spreadpandas.svg)](https://pypi.org/project/spreadpandas/)
[![PyPi](https://badgen.net/badge/icon/pypi?icon=pypi&label)](https://pypi.org/project/spreadpandas/)
## What is it?
SpreadPandas is a Python package to get the projected dimensions of a pandas data frame if it were to be transposed to a spreadsheet.

By creating a SpreadsheetMap object, understand which cells of a spreadsheet
the data frame would occupy, with detail for its header, index and body.

The goal is for this package to serve as a basis to perform more complex operations, like Excel customization or Google Sheet exporting, that require to establish a precise mapping between the data frame content and the target cells.

## Quick Start
### Installation
You can install EasyPred via `pip`
```
pip install spreadpandas
```
Alternatively, you can install EasyPred by cloning the project to your local directory
```
git clone https://github.com/FilippoPisello/SpreadPandas
```
And then run `setup.py`
```
python setup.py install
```

### Usage
Take an example pandas data frame:
```python
>>> import pandas as pd
>>> df = pd.DataFrame({"COL1" : [1, 2, 3], "COL2" : [4, 5, 6]})
>>> df
   COL1  COL2
0     1     4
1     2     5
2     3     6
```
Use SpreadMap to get its spreadsheet mapping. Suppose the index should not be brought over to the spreadsheet:
```python
>>> from spreadpandas import SpreadMap
>>> spreadmap = SpreadMap(df, keep_index=False)
>>> spreadmap.header.cells
('A1', 'B1')
>>> spreadmap.body.cells
('A2', 'B2', 'A3', 'B3', 'A4', 'B4')
>>> spreadmap.index
None
```
If the index should be kept instead:
```python
>>> spreadmap = SpreadMap(df, keep_index=True)
>>> spreadmap.header.cells
('B1', 'C1')
>>> spreadmap.body.cells
('B2', 'C2', 'B3', 'C3', 'B4', 'C4')
>>> spreadmap.index.cells
('A2', 'A3', 'A4')
```
The content can also be moved by skipping rows and/or columns:
```python
>>> spreadmap = SpreadMap(df, keep_index=True, skip_rows=2, skip_columns=1)
>>> spreadmap.header.cells
('C3', 'D3')
>>> spreadmap.body.cells
('C4', 'D4', 'C5', 'D5', 'C6', 'D6')
>>> spreadmap.index.cells
('B4', 'B5', 'B6')
```

## Dependencies
EasyPred depends only on `pandas`.

## Documentation
For the moment, this is the only documentation available.

## License
[MIT](LICENSE)