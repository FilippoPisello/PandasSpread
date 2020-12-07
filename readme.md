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
- [4. Roadmap for future developments](#4-roadmap-for-future-developments)

# 1. Overview
The Spreadsheet class is meant to identify which cells of a spreadsheet would be occupied by a pandas data frame based on its dimensions. The correspondence is carried out for the whole data frame and its sub elements, like the header, index and body.

It follows a simple example. Consider the 1x2 pandas data frame generated from the dictionary {"A" : [1], "B" : [2]}. By populating the sheet from the top left and keeping the index, the table would occupy the cells ["B1", "C1", "A2", "B2", "C2"]. In particular, ["B1", "C1"] would be the header, ["A2"] the index and ["B2", "C2"] the body.

## 1.1. Goal of the project
The Spreadsheet class should serve as a shared base for the modules to export pandas data frames to the different spreadsheet softwares, namely Excel and Google Sheet. These modules are then to be developed as sub-classes of the Spreadsheet class.

The logic behind this choice is that a pandas data frame occupies the same cells in a spreadsheet regardless of what the spreadsheet software is.

## 1.2. Code structure
The code is designed to covey a hierarchical division of the class' methods, through a system of tiers. These tiers are introduced by comment blocks which are made of two compact lines of "#" having a number in between. The number represent the tier.

The tiers are structured as follows:
- **Tier 1.1**: class properties
- **Tier 1.2**: class methods which are meant to be accessed by the user
- **Tier 2**: methods used to built tier 1 methods
- **Tier 3**: methods used to built tier 2 methods
- ...and so on
- **Tier D**: methods useful for tests or debugging

This structure should hopefully help the reader to understand how the simple pieces are assembled to construct more complex items. The individual methods are designed to follow as closely as possible the single-responsibility principle.

## 1.3. Terminology
Some terms should be unambiguously defined. Concerning the spreadsheet table elements:
- **Header**: the set of cells containing the labels of the columns.
- **Index**: the set of cells containing the labels of the rows.
- **Body**: the set of cells containing the data.
- **Table**: the union of the three set of cells above.

Concerning columns and rows:
- **Column letter/Column label**: the letter(s) used in the spreadsheet system to uniquely identify the columns. The most leftward column is labeled with "A", followed by "B" and so on. After "Z" it follows "AA", then "AB" and so on. After "ZZ" it follows "AAA", then "AAB" and so on.
- **Column index**: the position in space of a column. Following python conventions, the indexing is carried out from left to right, starting from 0. Thus the most leftward column has index 0.
- **Row number**: the number used in the spreadsheet system to uniquely identify the rows. The top row has number 1.

# 2. Required packages
This class relies on the following built-in package:
- **string**

And on the following additional packages:
- **numpy**
- **pandas**

# 3. Class elements
## 3.1. Arguments
The class intakes five arguments:
- **dataframe** : pandas data frame object (mandatory)
  - The pandas data frame to be considered.
- **keep_index** : Bool
  - If True, it is taken into account that the first column of the spreadsheet will be occupied by the index. All the dimensions will be adjusted as consequence.
- **skip_rows**: int
  - The number of rows which should be left empty at the top of the spreadsheet. Referring to excel row numbering, the table content starts at skip_rows + 1. If 0, content starts at row 1.
- **skip_cols**: int
  - The number of columns which should be left empty at the left of the spreadsheet. Referring to excel column labelling, the table content starts at letter with index skip_cols. If 0, content starts at column "A".
- **correct_lists**: Bool
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
- **self.table_coordinates**: [[int, int], [int, int]]
  - The coordinates of the cells at the top left and bottom right of the table, described by a list containing two lists of int of lent two. Each of the sub-lists identifies a spreadsheet cell. The first number is the column index, while the second one is the row number.
- **self.header_cells**: [str, ...]
  - A list of str where each str is a cell belonging to the header. Each cell is a pair formed by the column letter and the row number, like "A1".
- **self.index_cells**: [str, ...]
  - A list of str where each str is a cell belonging to the index. Each cell is a pair formed by the column letter and the row number, like "A1".
- **self.body_cell**: [str, ...]
  - A list of str where each str is a cell belonging to the body. Each cell is a pair formed by the column letter and the row number, like "A1".
- **self.table_cells**: [str, ...]
  - A list of str where each str is a cell belonging to the table. Each cell is a pair formed by the column letter and the row number, like "A1".

These can be found in tier 1.1 in the code base.

## 3.4. Methods
Coming soon

# 4. Roadmap for future developments
The following features are planned to be added soon to the class:
- **Select column method**:
  - This method should allow the user to insert the key of a column and obtain the set of cells belonging to it. The key parameter should accept different formats. In this sense, it is used as a reference the _"usecols"_ parameter of the data frame method read_excel from pandas.
- **Select row method**:
  - This method should be the equivalent of the one described above, for rows.

The developments mentioned now are instead potential evolutions of the class which are not planned to be implemented in the immediate future:
- **Clarification of the index role:**
  - At the moment, the behavior of the index can be considered unclear in case the user chooses _"False"_ for the class argument _"keep_index"_. What happens is that the real column index (in pandas _"df.index"_) is ignored, as it is supposed that it will not be included in the spreadsheet. Instead, _self.index_ will return the cells of the first column of the data frame. This is because it often happens in databases that the first column of the data frame contains the record keys.
