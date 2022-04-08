# SpeadPandas
The Spreadsheet class is meant to identify which cells of a spreadsheet would be occupied by a pandas data frame based on its dimensions. The correspondence is carried out for the whole data frame and its sub elements, like the header, index and body.

It follows a simple example. Consider the following table to be a 1x2 pandas data frame:
|     | Name   | Type                 |
| --- | ------ | -------------------- |
| 1   | Python | Programming Language |
| 2   | Word   | Text Editor          |

By populating the sheet from the top left and keeping the index, the table would occupy the cells ["B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"]. In particular, ["B1", "C1"] would be the header, ["A2", "A3"] the index and ["B2", "C2", "B3", "C3"] the body.

### Under construction...