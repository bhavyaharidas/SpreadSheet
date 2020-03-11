//Class represents each cell in the grid
class TableCell {
  constructor(id, displayValue, actualValue) {
    this.id = id; //cell id
    this.displayValue = displayValue; //Number
    this.actualValue = actualValue; //Formula if exists
    this.subscriptionObj = []; //Subscription objects corresponding to the table cell if exists
  }
}

let defaultRowCount = 15; // Number of rows on page load
let defaultColCount = 15; // Number of cols on page load
const SPREADSHEET_DB = "spreadsheet_db";

//Current indices
let selectedRowIndex = -1;
let selectedColIndex = -1;
let index;

//Global array holding grid tablecell values
var data = [];

//Order to determine precedence in math calc
const precedence = [
    ["*", "/"],
    ["+", "-"]
  ];

  const operators = {
    "+": function(a, b) {
      return a + b;
    },
    "-": function(a, b) {
      return a - b;
    },
    "*": function(a, b) {
      return a * b;
    },
    "/": function(a, b) {
      return a / b;
    }
  };

//Initialize empty grid upon page load
initializeGrid = () => {
  const data = [];
  for (let i = 0; i <= defaultRowCount; i++) {
    const child = [];
    for (let j = 0; j <= defaultColCount; j++) {
      child.push("");
    }
    data.push(child);
  }
  return data;
};

//Returns the global data if it is not empty, otherwise initializes data.
getData = () => {
  // let data = localStorage.getItem(SPREADSHEET_DB);
  if (data === undefined || data === null || data.length === 0) {
    return initializeGrid();
  }
  return data;
};

//Helper method to replace subscription objects while saving to local storage.
function jsonReplacer(key,value)
{
    if (key=="subscriptionObj") return undefined;
    else return value;
}

saveData = data => {
  localStorage.setItem(SPREADSHEET_DB, JSON.stringify(data, jsonReplacer));
};

//Creates the header row of the grid
createHeaderRow = () => {
  const tr = document.createElement("tr");
  tr.setAttribute("id", "h-0");
  for (let i = 0; i <= defaultColCount; i++) {
    const th = document.createElement("th");
    th.setAttribute("id", `h-0-${i}`);
    th.setAttribute("class", `${i === 0 ? "" : "column-header"}`);
    // th.innerHTML = i === 0 ? `` : `Col ${i}`;
    if (i !== 0) {
      const span = document.createElement("span");
      //Calculates the column header letter in upper case
      var res = String.fromCharCode(64 + i);
      span.innerHTML = res;
      span.setAttribute("class", "column-header-span"); 
      th.appendChild(span);
    }
    tr.appendChild(th);
  }
  return tr;
};

// Fill Data in created table from global data variable
populateTable = () => {
  const data = this.getData();
  if (data === undefined || data === null) return;
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    for (let j = 1; j < data[i].length; j++) {
      const cell = document.getElementById(`r-${i}-${j}`);
      cell.innerHTML = data[i][j].displayValue;
    }
  }
};

//Creates each table body row
createRow = rowNum => {
  const tr = document.createElement("tr");
  data = this.getData();
  tr.setAttribute("id", `r-${rowNum}`);
  for (let i = 0; i <= defaultColCount; i++) {
    const cell = document.createElement(`${i === 0 ? "th" : "td"}`);
    if (i === 0) {
      cell.contentEditable = false;
      cell.innerHTML = rowNum;
      cell.setAttribute("class", "row-header");
    } else if (!data[rowNum][i]) {
      let cellId = `r-${rowNum}-${i}`;
      //Creates new instance of TableCell and adds it to the data grid
      data[rowNum][i] = new TableCell(cellId, "", "");
      cell.contentEditable = true;
    } else {
      cell.contentEditable = true;
    }
    cell.setAttribute("id", `r-${rowNum}-${i}`);
    tr.appendChild(cell);
  }
  return tr;
};

createTableBody = tableBody => {
  for (let rowNum = 1; rowNum <= defaultRowCount; rowNum++) {
    tableBody.appendChild(this.createRow(rowNum));
  }
};

// Helper method to add row
//Adds onw row below the selected row
addRow = () => {
  if(selectedRowIndex > 0){
    let data = this.getData();
    let newRowIndex = selectedRowIndex + 1;
    const colCount = data[0].length;
    const newRow = new Array(colCount).fill("");
    data.splice(newRowIndex, 0, newRow);
    defaultRowCount++;
    // saveData(data);
    this.createSpreadsheet();
    //Reset the selected row index
    selectedRowIndex = -1;
  }
};

// Helper method to delete row
//Deletes the selected row
deleteRow = () => {
  if (selectedRowIndex > 0 && defaultRowCount > 1) {
    let data = this.getData();
    data.splice(selectedRowIndex, 1);
    defaultRowCount--;
    // saveData(data);
    this.createSpreadsheet();
    //Reset the selected row index
    selectedRowIndex = -1;
  }
};

// Helper method to add column
//Adds onw row below the selected column
addColumn = () => {
  let newColIndex = selectedColIndex + 1;
  if (newColIndex < 27 && newColIndex > 0 && defaultColCount < 26) {
    let data = this.getData();
    for (let i = 0; i <= defaultRowCount; i++) {
      data[i].splice(newColIndex, 0, "");
    }
    defaultColCount++;
    //  saveData(data);
    this.createSpreadsheet();
    selectedColIndex = -1;
  }
};

// Helper method to delete column
//Deletes the selected column
deleteColumn = currentCol => {
  let colPos = selectedColIndex;
  if (colPos < 27 && colPos > 0 && defaultColCount > 1) {
    let data = this.getData();
    for (let i = 0; i <= defaultRowCount; i++) {
      data[i].splice(colPos, 1);
    }
    defaultColCount--;
    //  saveData(data);
    this.createSpreadsheet();
    selectedColIndex = -1;
  }
};

// Helper method to import grid data from csv via upload
importFromCsv = () => {
  var fileUpload = document.getElementById("fileUpload");
        //File validation
        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.csv|.txt)$/;
        if (regex.test(fileUpload.value.toLowerCase())) {
            if (typeof (FileReader) != "undefined") {
                var reader = new FileReader();
                reader.onload = function (e) {
                    data = [];
                    //Holds the list of formulae in the csv file if any
                    let formulae = [];
                    var rows = e.target.result.split("\n");
                    if(rows[rows.length - 1] == ""){
                       rows.splice(rows.length - 1, 1);
                    }
                    //Create first row of empty data for header
                    let firstRow = [];
                    let columnCount = rows[0].split(",").length <= 26 ? rows[0].split(",").length : 26;
                    for(let i = 0; i <= columnCount; i++){
                      firstRow.push("");
                    }
                    data.push(firstRow);
                    //First create the skelton ignoring the formulae
                    for (let i = 0; i < rows.length; i++) {
                      let child = [];
                      //Create first column of empty data for row number
                      child.push("");
                      let row = rows[i].split(",");
                      for(let j = 0; j < columnCount; j++){
                        let cellId = `r-${i + 1}-${j + 1}`;
                        let displayValue = ""
                        //If cell data not a formula 
                        if(!row[j].startsWith("=")){
                          displayValue = row[j];
                        }else{ //If cell data is a formular add it to formulae array.
                          let formula = [];
                          formula.push(i);
                          formula.push(j);
                          formula.push(row[j]);
                          formula.push(cellId);
                          formulae.push(formula);
                        }
                        let cell = new TableCell(cellId, displayValue, "");
                        child.push(cell);
                      }
                      data.push(child);
                    }
                    //saveData(data);
                    createSpreadsheet();

                    //then calculate formulae, create and subscribe observables
                    for(let i = 0; i < formulae.length; i++){
                      let formulaCell = document.getElementById(formulae[i][3]);
                      formulaCell.innerHTML = formulae[i][2];
                      //Call the helper method to mimic a focusout
                      elementFocusout(formulaCell);
                }
              }
                reader.readAsText(fileUpload.files[0]);
            } else {
                alert("This browser does not support HTML5.");
            }
        } else {
            alert("Please upload a valid CSV file.");
        }
}

//Helper method to download grid as csv
downloadCSV = (csv, filename) => {
  var csvFile;
  var downloadLink;

  // CSV file
  csvFile = new Blob([csv], {type: "text/csv"});

  // Download link
  downloadLink = document.createElement("a");
  downloadLink.download = filename;

  // Create a link to the file
  downloadLink.href = window.URL.createObjectURL(csvFile);

  // Hide download link
  downloadLink.style.display = "none";
  document.body.appendChild(downloadLink);

  // Click download link
  downloadLink.click();
}

//Helper method to convert global data variable to csv grid
exportToCsv = () => {
  var filename = 'grid.csv'
  var csv = [];
  const spreadsheetData = this.getData();

  for (var i = 1; i < spreadsheetData.length; i++) {
    var row = []
    var col = spreadsheetData[i];
    
    //if a formula exists, then add it. otherwise add the display number
    for (var j = 1; j < col.length; j++){
      if(col[j].actualValue !== "")
        row.push(col[j].actualValue);
      else
        row.push(col[j].displayValue);
    }  
    csv.push(row.join(","));     
  }
  // Download CSV file
  downloadCSV(csv.join("\n"), filename);
}

//Creates various controls for the buttons
//Adds event listeners associated with its fucntionality.
createControls = () => {
  const addRowBtn = document.getElementById("addRow");
  const addColumnBtn = document.getElementById("addColumn");
  const deleteRowBtn = document.getElementById("deleteRow");
  const deleteColumnBtn = document.getElementById("deleteColumn");
  const importCsvBtn = document.getElementById("importCsv");
  const exportCsvBtn = document.getElementById("exportCsv");

  addRowBtn.addEventListener("click", function(e) {
    if (e.target) {
        const indices = index.split("-");
        addRow();  
    }
  });

  addColumnBtn.addEventListener("click", function(e) {
    if (e.target) {
        const indices = index.split("-");
        addColumn();  
    }
  });

  deleteRowBtn.addEventListener("click", function(e) {
    if (e.target) {
        const indices = index.split("-");
        deleteRow();  
    }
  });

  deleteColumnBtn.addEventListener("click", function(e) {
    if (e.target) {
        const indices = index.split("-");
        deleteColumn();  
    }
  });

  importCsvBtn.addEventListener("click", function(e) {
    importFromCsv();
    document.getElementById("fileUpload").value = "";
  });

  exportCsvBtn.addEventListener("click", function() {
    exportToCsv();
  });
  
}

//Helper method to calculate the given math formula
calculateExpression = formula => {
  let formulaArr = [];
  let input = [];
  //If the formula starts with sum, add all cell values in the range with a + sign btw them
  if(formula.startsWith("=SUM")){
    formula = formula.replace("=SUM(","");
    formula = formula.replace(")","");
    formulaArr = formula.split(":");
    //Same Column
    if(formulaArr[0].charAt(0) === formulaArr[1].charAt(0)){ 
      let column = formulaArr[0].charCodeAt(0);
      let startIndex = parseInt(formulaArr[0].charAt(1));
      let endIndex = parseInt(formulaArr[1].charAt(1));
      for(let i = startIndex; i <= endIndex; i++){
        let cellId = `r-${i}-${column - 64}`
        input.push(document.getElementById(cellId).innerHTML);
        if(i !== endIndex){
          input.push("+");
        }
      }
    }
    //Same Row
    else{
      let row = formulaArr[0].charAt(1);
      let startIndex = parseInt(formulaArr[0].charCodeAt(0)) - 64;
      let endIndex = parseInt(formulaArr[1].charCodeAt(0)) - 64;
      for(let i = startIndex; i <= endIndex; i++){
        let cellId = `r-${row}-${i}`
        input.push(document.getElementById(cellId).innerHTML);
        if(i !== endIndex){
          input.push("+");
        }
      }
    }
  } //if formula does not start with SUM, add all cell data and operators to the input array.
  else{
    formulaArr = formula.split(/([=*/%+-])/g).splice(2);
    formulaArr.forEach(operand => {
      if (operand.toLowerCase() !== operand.toUpperCase()) {
          let cellId = `r-${operand.charAt(1)}-${operand.charCodeAt(0) - 64}`
          input.push(document.getElementById(cellId).innerHTML);
      } else {
        input.push(operand);
      }
    });
  }
  //Continue the process till the formula array is empty
  while (input.length > 1) {
    // find the first operator at the lowest level
    let reduceAt = 0;
    let found = false;
    let newInput = [];
    for (let i = 0; i < precedence.length; i++) {
      for (let j = 1; j < input.length - 1; j++) {
        if (precedence[i].indexOf(input[j]) >= 0) {
          reduceAt = j;
          found = true;
          break;
        }
      }
      if (found) break;
    }

    // if we didn't find one, bail
    if (!found) return;

    // otherwise, reduce that operator

    var f = operators[input[reduceAt]];

    for (let i = 0; i < reduceAt - 1; i++) {
      newInput.push(input[i]);
    }
    newInput.push(
      "" + f(parseFloat(input[reduceAt - 1]), parseFloat(input[reduceAt + 1]))
    );
    for (let i = reduceAt + 2; i < input.length; i++) {
      newInput.push(input[i]);
    }
    input = newInput;
  }
  return input[0];
};

//Makes use of fromEvent mehod of rxjs to return and obeservable associated with the given cell.
createCellObservable = cellId => {
  return rxjs.fromEvent(document.getElementById(cellId), "focusout");
};


createObservables = formula => {
  let formulaArr = [];
  let currentObs = [];
  if(formula.startsWith("=SUM")){
    formula = formula.replace("=SUM(","");
    formula = formula.replace(")","");
    formulaArr = formula.split(":");
    //Same Column
    if(formulaArr[0].charAt(0) === formulaArr[1].charAt(0)){ 
      let column = formulaArr[0].charCodeAt(0);
      let startIndex = parseInt(formulaArr[0].charAt(1));
      let endIndex = parseInt(formulaArr[1].charAt(1));
      for(let i = startIndex; i <= endIndex; i++){
        let cellId = `r-${i}-${column - 64}`
        currentObs.push(
          this.createCellObservable(cellId)
        );
      }
    }
    //Same Row
    else{
      let row = formulaArr[0].charAt(1);
      let startIndex = parseInt(formulaArr[0].charCodeAt(0)) - 64;
      let endIndex = parseInt(formulaArr[1].charCodeAt(0)) - 64;
      for(let i = startIndex; i <= endIndex; i++){
        let cellId = `r-${row}-${i}`
        currentObs.push(
          this.createCellObservable(cellId)
        );
      }
    }
  }
  else{
    formulaArr = formula.split(/([=*/%+-])/g);
    if (formulaArr.length > 2) {
      for (let i = 1; i < formulaArr.length; i++) {
        if (formulaArr[i].length === 2) {
          let cellId = `r-${formulaArr[i].charAt(1)}-${formulaArr[i].charCodeAt(0) - 64}`;
          currentObs.push(
            this.createCellObservable(cellId)
          );
        }
      }
    }
  }
  return currentObs;
};

createSpreadsheet = () => {
  const spreadsheetData = this.getData();
  defaultRowCount = spreadsheetData.length - 1 || defaultRowCount;
  defaultColCount = spreadsheetData[0].length - 1 || defaultColCount;

  const tableHeaderElement = document.getElementById("table-headers");
  const tableBodyElement = document.getElementById("table-body");

  const tableBody = tableBodyElement.cloneNode(true);
  tableBodyElement.parentNode.replaceChild(tableBody, tableBodyElement);
  const tableHeaders = tableHeaderElement.cloneNode(true);
  tableHeaderElement.parentNode.replaceChild(tableHeaders, tableHeaderElement);

  tableHeaders.innerHTML = "";
  tableBody.innerHTML = "";

  tableHeaders.appendChild(createHeaderRow(defaultColCount));
  createTableBody(tableBody, defaultRowCount, defaultColCount);

  populateTable();

  /*
  tableBody.addEventListener("focusin", function(e) {
    if (e.target && e.target.nodeName === "TD") {
      let item = e.target;
      const indices = item.id.split("-");
      item.innerHTML = data[indices[1]][indices[2]].actualValue;
    }
  });
  */

  elementFocusout = (item) => {
    
      const indices = item.id.split("-");
      let data = getData();
      let currentCellData = item.innerHTML;
      if(currentCellData.startsWith("=")){
        let calcValue = calculateExpression(currentCellData);
        data[indices[1]][indices[2]].actualValue = item.innerHTML;
        data[indices[1]][indices[2]].displayValue = calcValue;
        document.getElementById(item.id).innerHTML = calcValue;
        createObservables(
            data[indices[1]][indices[2]].actualValue
        ).forEach(observable => {
          data[indices[1]][indices[2]].subscriptionObj.push(observable.subscribe(() => {
            document.getElementById(item.id).innerHTML = calculateExpression(
                data[indices[1]][indices[2]].actualValue
            );
            data[indices[1]][indices[2]].displayValue = document.getElementById(item.id).innerHTML;
          }));
        });
      }else{
        if(currentCellData !== data[indices[1]][indices[2]].displayValue && data[indices[1]][indices[2]].actualValue.startsWith("=")){
          data[indices[1]][indices[2]].subscriptionObj.forEach(subObj => {
            subObj.unsubscribe();
          })
          data[indices[1]][indices[2]].actualValue = item.innerHTML;
        }
        //data[indices[1]][indices[2]].actualValue = item.innerHTML;
        data[indices[1]][indices[2]].displayValue = item.innerHTML;
        //saveData(data);
      }
  }

  // attach focusout event listener to whole table body container
  tableBody.addEventListener("focusout", function(e) {
    if (e.target && e.target.nodeName === "TD") {
      let item = e.target;
      elementFocusout(item);
    }
  });

  // Attach click event listener to table body
  tableBody.addEventListener("click", function(e) {
    clearSelection();
    if (e.target) {
      index = e.target.id;
      if (e.target.className === "row-header") {
        e.target.classList.add("selected");
        e.target.parentNode.classList.add("selected");
        selectedRowIndex = parseInt(e.target.parentNode.id.split("-")[1]);
      }
      if (e.target.className === "row-insert-top") {
        const indices = e.target.parentNode.id.split("-");
        addRow(parseInt(indices[2]), "top");
      }
      if (e.target.className === "row-insert-bottom") {
        const indices = e.target.parentNode.id.split("-");
        addRow(parseInt(indices[2]), "bottom");
      }
      if (e.target.className === "row-delete") {
        const indices = e.target.parentNode.id.split("-");
        deleteRow(parseInt(indices[2]));
      }
    }
  });

  // Helper function to highlight the selected headers
  highlightColumn = colId => {
      let data = this.getData();
      for (let i = 1; i < data.length; i++) {
        document.getElementById(`r-${i}-${colId}`).classList.add("selected");
      }
  };

    // Reset the previously selected headers
    clearSelection = () => {
      selectedRowIndex = -1;
      selectedColIndex = -1;
      document.querySelectorAll(".selected").forEach(node => {
        node.classList.remove("selected");
      });
    };

  // Attach click event listener to table headers
  tableHeaders.addEventListener("click", function(e) {
    clearSelection();
    if (e.target && e.target.className === "column-header") {
      e.target.classList.add("selected");
      highlightColumn(parseInt(e.target.id.split("-")[2]));
      selectedColIndex = parseInt(e.target.id.split("-")[2]);
    }
  });
};

//Call base methods for initial page load
createSpreadsheet();
createControls();
