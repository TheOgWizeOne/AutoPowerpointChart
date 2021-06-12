

help = """
Transfer chart data from a specially formatted excel workbook

Default values for file locations are loaded from config.txt file.

Start transfer - > Use "Press here to start file transfer button" 
or select Run>Transfer from menus 

To show current files and parameters for transfer use "Toggle show file"

The main defaults are in grey - press the file name button to select an alternative.
Template presentation - this is the doner powerpoint file.
used to create the presentation. This sets the look and feel of the output file.
Excel data file - the excel workbook stores the chart data in a specific format (below).

Output presentation is the filename of the destination powerpoint - this file is overwritten.

The darker slate grey boxes show the template files for each chart type.
These can be changed using the relevant button and selecting a file.

If you would like to save your selections use the "Save defaults" option from the file menu.
This saves the current file options as default and will be loaded next time.

Excel format:
An "index" sheet - with a 4 col table starting at A1. Each chart to be
transfered should be a row of this table with name,sheets,type,label data
The "sheets" col refers the sheet with the chart data, the type is the chart type
(supported types: col_clustered, col_stacked, col_stacked_100, bar_clustered, 
bar_stacked, bar_stacked_100, pie, and line) and the label col sets the label for the
x-axis data as either "per" for precentages  or "num" for count or number variables.

Each sheet with chart data should be located from A1 with the x-axis categories
 in the first column label "category"

"""


about = """MIT License

Copyright (c) 2021 Market Prescience Ltd (www.marketprescience.com)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""