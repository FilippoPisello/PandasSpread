# Table of Contents <!-- omit in toc -->
- [1. Overview](#1-overview)
  - [1.1. Goal of the project](#11-goal-of-the-project)
  - [1.2. Code structure](#12-code-structure)
  - [1.3. Terminology](#13-terminology)
- [2. Required packages](#2-required-packages)
- [3. Class elements](#3-class-elements)
  - [3.1. Arguments](#31-arguments)
  - [3.2. Attributes](#32-attributes)
  - [3.3. Properties](#33-properties)
  - [3.4. Methods](#34-methods)
    - [3.4.1. column()](#341-column)
    - [3.4.2. row()](#342-row)
- [4. SpreadsheetElement class](#4-spreadsheetelement-class)
  - [4.1. Arguments](#41-arguments)
  - [4.2. Properties](#42-properties)
- [5. Roadmap for future developments](#5-roadmap-for-future-developments)
  - [5.1. Planned developments](#51-planned-developments)
  - [5.2. Potential developments](#52-potential-developments)

# 1. Overview
The Spreadsheet class is meant to identify which cells of a spreadsheet would be occupied by a pandas data frame based on its dimensions. The correspondence is carried out for the whole data frame and its sub elements, like the header, index and body.

It follows a simple example. Consider the following table to be a 1x2 pandas data frame:
|     | Name   | Type                 |
| --- | ------ | -------------------- |
| 1   | Python | Programming Language |
| 2   | Word   | Text Editor          |

By populating the sheet from the top left and keeping the index, the table would occupy the cells ["B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"]. In particular, ["B1", "C1"] would be the header, ["A2", "A3"] the index and ["B2", "C2", "B3", "C3"] the body.

## 1.1. Goal of the project
The Spreadsheet class should serve as a shared base for the modules to export pandas data frames to the different spreadsheet softwares, namely Excel and Google Sheet. These modules are then to be developed as sub-classes of the Spreadsheet class.

The logic behind this choice is that a pandas data frame occupies the same cells in a spreadsheet regardless of what the spreadsheet software is.

## 1.2. Code structure
The project's code is contained in two files:
- **spreadsheet.py**: the main file which contains the Spreadsheet class. Except for point 4, all the indications in this guide refer to this file and class.
- **spreadsheet_element.py**: file which contains the support class SpredsheetElement. More on this class at point 4.

The code is designed to convey a hierarchical division of the class' methods. Different code portions are introduced by comment blocks which are made of two compact lines of "#" having a number in between.

The sections are structured as follows:
- **Part 1**: elements with external scope
  - **Part 1.1**: class properties
    - **Part 1.1.1**: subparts' coordinates
    - **Part 1.1.2**: subparts' objects
  - **Part 1.2**: main class methods. These are meant to be accessed by the user
- **Part 2**: worker methods
  - **Part 2.1**: methods used in attributes
  - **Part 2.2**: chain of methods used for the row/column methods in 1.2
  - **Part 2.3**: methods used as building blocks in multiple parts of the class
- **Tier D**: methods useful for tests or debugging

The methods' are ordered so that if method B is invoked by method A, then B will be below A. This structure should hopefully help the reader to understand how the simple pieces are assembled to construct more complex items.

The individual methods are designed to follow as closely as possible the single-responsibility principle.

## 1.3. Terminology
Some terms should be unambiguously defined. Concerning the spreadsheet table elements:
- **Header**: the set of cells containing the labels of the columns.
- **Index**: the set of cells containing the labels of the rows.
- **Body**: the set of cells containing the data.
- **Table**: the union of the three set of cells above.

Concerning columns and rows:
- **Column letter/Column label**: the letter(s) used in the spreadsheet system to uniquely identify the columns. The most leftward column is labeled with "A", followed by "B" and so on. After "Z" it follows "AA", then "AB" and so on. After "ZZ" it follows "AAA", then "AAB" and so on.
- **Column index**: the position in space of a column. Following python conventions, the indexing is carried out from left to right, starting from 0. Thus the most leftward column has index 0. It follows that column "A" is equal to column 0.
- **Row number**: the number used in the spreadsheet system to uniquely identify the rows. The top row has number 1.

# 2. Required packages
This class relies on the following built-in packages:
- **string**
- **typing**

And on the following additional packages:
- **numpy**
- **pandas**

# 3. Class elements
## 3.1. Arguments
The class intakes five arguments:
- **dataframe** : pandas data frame object (mandatory)
  - The pandas data frame to be considered.
- **index** : Bool, default=False
  - If True, it is taken into account that the first column of the spreadsheet will be occupied by the index. All the dimensions will be adjusted as consequence.
- **starting_cell**: str, default="A1"
  - The cell that represents the top left corner of the data frame in the
  spreadsheet. No cells above or at the left of this cell will be mapped.
- **correct_lists**: Bool, default=False
  - If True, the lists stored as the data frame entries are modified to be more readable in the traditional spreadsheet softwares. This happens in four ways. (1) Empty lists are replaced by missing values. (2) Missing values are removed from within the lists. (3) Lists of len 1 are replaced by the single element they contain. (4) Lists are replaced by str formed by their elements separated by commas.

## 3.2. Attributes
The spreadsheet object has four attributes, which correspond to the arguments:
- **self.df** : pandas data frame object
- **self.keep_index** : Bool
- **self.skip_rows**: int
- **self.skip_cols**: int

## 3.3. Properties
The spreadsheet object has eight properties:
- **self.indexes_depth** : [int, int]
  - A list of len two containing the number of multilevels respectively for  the index and the columns of the pandas data frame used as input. If the index is not a multiindex, the first number will be 1. The same is true for the columns.
- **self.header_coordinates** : [[int, int], [int, int]]
  - The coordinates of the cells at the top left and bottom right of the header, described by a list containing two lists of int of lent two. Each of the sub-lists identifies a spreadsheet cell. The first number is the column index, while the second one is the row number.
- **self.index_coordinates**: [[int, int], [int, int]]
  - The coordinates of the cells at the top left and bottom right of the index, described by a list containing two lists of int of lent two. Each of the sub-lists identifies a spreadsheet cell. The first number is the column index, while the second one is the row number.
- **self.body_coordinates**: [[int, int], [int, int]]
  - The coordinates of the cells at the top left and bottom right of the body, described by a list containing two lists of int of lent two. Each of the sub-lists identifies a spreadsheet cell. The first number is the column index, while the second one is the row number.
- **self.header**: SpreadsheetElement object
  - An object which allows to access various header's properties. More on these objects at point 4.
- **self.index**: SpreadsheetElement object
  - An object which allows to access various index's properties. More on these objects at point 4.
- **self.body**: SpreadsheetElement object
  - An object which allows to access various body's properties. More on these objects at point 4.
- **self.table**: SpreadsheetElement object
  - An object which allows to access various table's properties. More on these objects at point 4.

These can be found in tier 1.1 in the code base.

## 3.4. Methods
This section just includes the methods which are meant to be accessed by the user, thus of tier 1.2. For further info on the methods of lower tiers, please consult directly their docstrings.

### 3.4.1. column()
Returns the list of cells making up a column in the form ["A1", "A2", ...].

The column can be identified using its numeric index, its spreadsheet letter or its label. Multiple columns can be accessed as well.

Note that when evaluating a bit of a str key, the match with the spreadsheet letters is carried out if no match is found with column names. This implies that if "A" is passed and a column named "A" exists, then the program will match the key with that column, regardless of its spreadsheet letter. In cases where one wants to refer to columns named as spreadsheet letters, index or column label should be used as key.

**Arguments**
- **key**: int, str, list, tuple

  It is used to identify the column(s) whose cells should be returned.
  - Single column index can be passed as int (ex: 1)
  - Multiple column index can be passed as list/tuple of int (ex: [1, 2])
  - Single spreadsheet column letter can be passed as str (ex: "A")
  - Multiple spreadsheet column letters can be passed as list/tuple of str (ex: ["A", "B"]). Alternatively, they can be passed as str being comma separated "A, B".
  - A range of spreadsheet columns can be passed as str with the first and last letter divided by a colon ("A:C"). This range is inclusive on both sides.
  - Commas and colons can be combined as well ("A,C:E")
  - The last three options apply in the same way to column labels. (ex: "Foo")(ex: ["Foo", "Bar"]) (ex: "Foo, Bar") (ex: "Foo:Bar")
- **include_header**: Bool, default=False

  It specifies whether the cells making up the column's header should be included or not. If False, it only returns the cells containing data for that column(s).

**Examples**

Consider the following table to be the pandas data frame passed in the constructor with all the parameters left to default values:
| Foo | Bar | Baz | Qux |
| --- | --- | --- | --- |
| 1   | 2   | 3   | 4   |
| 5   | 6   | 7   | 8   |

The object spreadsheet is created.
- spreadsheet.column("A", True) --> ["A1", "A2", "A3"]
- spreadsheet.column("Foo", True) --> ["A1", "A2", "A3"]
- spreadsheet.column(0, True) --> ["A1", "A2", "A3"]
- spreadsheet.column("A", False) --> ["A2", "A3"]
- spreadsheet.column("A, C", False) --> ["A2", "A3", "C2", "C3"]
- spreadsheet.column(["A", "C"], False) --> ["A2", "A3", "C2", "C3"]
- spreadsheet.column("A:C", False) --> ["A2", "A3", "B2", "B3", "C2", "C3"]
- spreadsheet.column("Foo:Baz", False) --> ["A2", "A3", "B2", "B3", "C2", "C3"]
- spreadsheet.column([0, 3], False) --> ["A2", "A3", "D2", "D3"]

### 3.4.2. row()
Returns the list of cells making up a row in the form ["A1", "B1", ...].

The row can be identified using its numeric index or its spreadsheet number, where row number = index + 1. It is possible to access multiple rows.

**Arguments**
- **key**: int, str, list, tuple

  It is used to identify the row(s) whose cells should be returned.
  - Single row index can be passed as int (ex: 1)
  - Multiple row index can be passed as list/tuple of int (ex: [1, 2])
  - Single spreadsheet row number can be passed as str (ex: "1")
  - Multiple spreadsheet row numbers can be passed as list/tuple of str (ex: ["1", "2"]). Alternatively, they can be passed as str being comma separated "1, 2"
  - A range of row numbers can be passed as str with the first and last index divided by a colon ("1:3"). This range is inclusive on both sides.
  - Commas and colons can be combined as well ("1,2:4")
- **include_index**: Bool, default=False

  It specifies whether the cells making up the column's index should be included or not. If False, it only returns the cells containing data for that row(s).

**Examples**

Consider the following table to be the pandas data frame passed in the constructor with all the parameters left to default values:
| Foo | Bar | Baz |
| --- | --- | --- |
| 1   | 2   | 3   |
| 5   | 6   | 7   |
| 8   | 9   | 10  |

The object spreadsheet is created.
- spreadsheet.row("2", True) --> ["A2", "B2", "C2", "D2"]
- spreadsheet.row(1, True) --> ["A2", "B2", "C2", "D2"]
- spreadsheet.row("2", False) --> ["A2", "B2", "C2"]
- spreadsheet.row("2, 3", False) --> ["A2", "B2", "C2", "A3", "B3", "C3"]
- spreadsheet.row(["2", "3"], False) --> ["A2", "B2", "C2", "A3", "B3", "C3"]
- spreadsheet.row("2:3", False) --> ["A2", "B2", "C2", "A3", "B3", "C3"]
- spreadsheet.row([1, 2], False) --> ["A2", "B2", "C2", "A3", "B3", "C3"]

# 4. SpreadsheetElement class
The SpreadSheet element class is a support class created to limit the code repetitions. It proved to be necessary as the number of representations for the various spreadsheet parts grew in number. This class can be found in the dedicated file "spreadsheet_element.py".

A SpreadsheetElement object describes a portion of a spreadsheet. In this class, an object of this type is created for each of the four table parts mentioned in the overview: header, index, body, table.
## 4.1. Arguments
The class intakes two arguments:
- coordinates: [[int, int], [int, int]], (mandatory)
  - A system of coordinates which identifies a rectangle of cells. It is a list containing two lists of length two, each containing two numbers which univocally identify a cell. The two numbers are respectively the row index and the column number. The two cells are the top left and bottom right one.
- style: None or Style object, default=None
  - An object of the adequate type can be passed depending on the spreadsheet subclass used. For CustomExcel, one should use and ExcelStyle object. If None, no action will be taken.

## 4.2. Properties
The class has two properties:
- cells: [str, ..., str]
  - The ist of cells making up the object in the form ["A1", "A2"]
- cells_range: str
  - A str in this form "A1:B3". Where "A1" is the top left cell of the object while "B3" is the bottom right.

# 5. Roadmap for future developments
## 5.1. Planned developments
There are no developments planned for the near future.

## 5.2. Potential developments
The developments mentioned now are instead potential evolutions of the class which are not planned to be implemented in the immediate future:
- **Clarification of the index role:**
  - At the moment, the behavior of the index can be considered unclear in case the user chooses _"False"_ for the class argument _"index"_. What happens is that the real column index (in pandas _"df.index"_) is ignored, as it is supposed that it will not be included in the spreadsheet. Instead, _self.index_ will return the cells of the first column of the data frame. This is because it often happens in databases that the first column of the data frame contains the record keys.