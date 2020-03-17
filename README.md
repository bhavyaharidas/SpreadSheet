# Spread Sheet Web App

## Web Design Project to learn about JavaScript events & RxJS

### Problem Statement - Create a spreadsheet application like Google sheets using Javascript, RxJS, and CSS.
Develop a spreadsheet application containing the following functionalities-
1. Add new row below the selected row.
2. Add new column to the right of the selected column.
3. Delete selected row.
4. Delete selected column.
5. Select multiple rows or columns and display their sum in a cell by using a formula.
6. Perform simple algebraic operations (+, -, *, /) with cell references.
7. Export the sheet as a CSV file.
8. Upload a CSV.

Name - Bhavya Haridas

### Technologies used -
1. HTML
2. CSS
3. Javascript
4. RxJs

### Requirements -
1. Web Browser

### Steps to Replicate -

1. Clone repository - git clone https://github.com/bhavyaharidas/SpreadSheet.git
2. Open folder SpreadSheet
3. Double click on index.html to open the web page on a browser

### Steps to make use of the functionalities.

1. Type your sample numerical data on the grid.
#### Add/Delete Row
2. Select any one row by clicking on the serial number on the leftmost non-data column.
3. Click on "Add Row" to see that a row has been added below the selected row.
4. Click on "Delete Row" to see that the selected row has been deleted.
#### Add/Delete Column
5. Select any one column by clicking on the letter on the header row.
6. Click on "Add Column" to see that a column has been added to the right of the selected column.
7. Click on "Delete Column" to see that the selected column has been deleted.
#### Mathematical Calculations
6. Enter your formula in any one of the cell similar to the following examples
    i.   =A1+A2+A3
    ii.  =SUM(A1:A3)
    iii. =A1+B1-C1*D1
7. Verify that the result has been displayed and also any change in the operand cells updates the result cell.
8. Clear off the data in the result cell to see that the binding has been cut off from the operand cells.
#### Export to CSV
9. Click on Export button which downloads a csv file named "grid" with the entered data (including the formulae) in it.
#### Import CSV
10. Choose the file to upload and click on Import to load any CSV data onto the grid.

### Libraries used -

RxJs -
https://unpkg.com/rxjs/bundles/rxjs.umd.js

### References -

1. https://developer.mozilla.org/en-US/docs/Web/JavaScript
2. https://rxjs-dev.firebaseapp.com/guide/overview