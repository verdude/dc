### Program Description

This software compares the first sheet in two xlsx files. A table of differences is constructed. Every cell in the first file is compared to the corresponding cell in the second file. If the value of the cells differ in some way other than casing (upper versus lower case), then the cells are displayed as a row in the table along with the cell coordinates. If the two sheets have a different number of rows, that is also reported to the user as a warning dialog.

### Technical Details

This software uses un-edited libraries from the [SheetJS](https://github.com/SheetJS/js-xlsx) project. Methods from the [example file](https://github.com/SheetJS/js-xlsx/blob/master/index.html) were adapted and used in the [index file](https://github.com/verdude/dc/blob/master/index.html) of this project.
