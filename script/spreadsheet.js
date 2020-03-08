let defaultRowCount = 15; // No of rows
let defaultColCount = 12; // No of cols
const SPREADSHEET_DB = "spreadsheet_db";
let index;
var data = [];

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

class TableCell {
  constructor(id, displayValue, actualValue) {
    this.id = id;
    this.displayValue = displayValue;
    this.actualValue = actualValue;
  }
}

initializeData = () => {
  // console.log("initializeData");
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

getData = () => {
  // let data = localStorage.getItem(SPREADSHEET_DB);
  if (data === undefined || data === null || data.length === 0) {
    return initializeData();
  }
  return data;
};

saveData = data => {
  localStorage.setItem(SPREADSHEET_DB, JSON.stringify(data));
};

resetData = data => {
  localStorage.removeItem(SPREADSHEET_DB);
  this.createSpreadsheet();
};

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
      var res = String.fromCharCode(64 + i);
      span.innerHTML = res;
      span.setAttribute("class", "column-header-span");
      const dropDownDiv = document.createElement("div");
      dropDownDiv.setAttribute("class", "dropdown");
      dropDownDiv.innerHTML = `<div id="col-dropdown-${i}" class="dropdown-content">
          <p class="col-insert-left">Insert 1 column left</p>
          <p class="col-insert-right">Insert 1 column right</p>
          <p class="col-delete">Delete column</p>
        </div>`;
      th.appendChild(span);
      th.appendChild(dropDownDiv);
    }
    tr.appendChild(th);
  }
  return tr;
};

createTableBodyRow = rowNum => {
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
    tableBody.appendChild(this.createTableBodyRow(rowNum));
  }
};

// Fill Data in created table from localstorage
populateTable = () => {
  const data = this.getData();
  if (data === undefined || data === null) return;
  for (let i = 1; i < data.length; i++) {
    for (let j = 1; j < data[i].length; j++) {
      const cell = document.getElementById(`r-${i}-${j}`);
      cell.innerHTML = data[i][j].displayValue;
    }
  }
};

// Utility function to add row
addRow = (currentRow, direction) => {
  let data = this.getData();
  const colCount = data[0].length;
  const newRow = new Array(colCount).fill("");
  if (direction === "top") {
    data.splice(currentRow, 0, newRow);
  } else if (direction === "bottom") {
    data.splice(currentRow + 1, 0, newRow);
  }
  defaultRowCount++;
  saveData(data);
  this.createSpreadsheet();
};

// Utility function to delete row
deleteRow = currentRow => {
  let data = this.getData();
  data.splice(currentRow, 1);
  defaultRowCount++;
  saveData(data);
  this.createSpreadsheet();
};

// Utility function to add columns
addColumn = (currentCol, direction) => {
  let data = this.getData();
  for (let i = 0; i <= defaultRowCount; i++) {
    if (direction === "left") {
      data[i].splice(currentCol, 0, "");
    } else if (direction === "right") {
      data[i].splice(currentCol + 1, 0, "");
    }
  }
  defaultColCount++;
  saveData(data);
  this.createSpreadsheet();
};

// Utility function to delete column
deleteColumn = currentCol => {
  let data = this.getData();
  for (let i = 0; i <= defaultRowCount; i++) {
    data[i].splice(currentCol, 1);
  }
  defaultColCount++;
  saveData(data);
  this.createSpreadsheet();
};

// Map for storing the sorting history of every column;
const sortingHistory = new Map();

// Utility function to sort columns
sortColumn = currentCol => {
  let spreadSheetData = this.getData();
  let data = spreadSheetData.slice(1);
  if (!data.some(a => a[currentCol] !== "")) return;
  if (sortingHistory.has(currentCol)) {
    const sortOrder = sortingHistory.get(currentCol);
    switch (sortOrder) {
      case "desc":
        data.sort(ascSort.bind(this, currentCol));
        sortingHistory.set(currentCol, "asc");
        break;
      case "asc":
        data.sort(dscSort.bind(this, currentCol));
        sortingHistory.set(currentCol, "desc");
        break;
    }
  } else {
    data.sort(ascSort.bind(this, currentCol));
    sortingHistory.set(currentCol, "asc");
  }
  data.splice(0, 0, new Array(data[0].length).fill(""));
  saveData(data);
  this.createSpreadsheet();
};

// Compare Functions for sorting - ascending
const ascSort = (currentCol, a, b) => {
  let _a = a[currentCol];
  let _b = b[currentCol];
  if (_a === "") return 1;
  if (_b === "") return -1;

  // Check for strings and numbers
  if (isNaN(_a) || isNaN(_b)) {
    _a = _a.toUpperCase();
    _b = _b.toUpperCase();
    if (_a < _b) return -1;
    if (_a > _b) return 1;
    return 0;
  }
  return _a - _b;
};

// Descending compare function
const dscSort = (currentCol, a, b) => {
  let _a = a[currentCol];
  let _b = b[currentCol];
  if (_a === "") return 1;
  if (_b === "") return -1;

  // Check for strings and numbers
  if (isNaN(_a) || isNaN(_b)) {
    _a = _a.toUpperCase();
    _b = _b.toUpperCase();
    if (_a < _b) return 1;
    if (_a > _b) return -1;
    return 0;
  }
  return _b - _a;
};

importFromCsv = () => {
  var fileUpload = document.getElementById("fileUpload");
        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.csv|.txt)$/;
        if (regex.test(fileUpload.value.toLowerCase())) {
            if (typeof (FileReader) != "undefined") {
                var reader = new FileReader();
                reader.onload = function (e) {
                    var table = document.createElement("table");
                    var rows = e.target.result.split("\n");
                    const data = [];
                    for (let i = 0; i < rows.length; i++) {
                      let child = [];
                      let row = rows[i].split(",");
                      for(let j = 0; j < row.length; j++){
                        child.push(row[j]);
                      }
                      data.push(child);
                    }
                    saveData(data);
                    createSpreadsheet();
                }
                reader.readAsText(fileUpload.files[0]);
            } else {
                alert("This browser does not support HTML5.");
            }
        } else {
            alert("Please upload a valid CSV file.");
        }
}

downloadCSV = (csv, filename) => {
  var csvFile;
  var downloadLink;

  // CSV file
  csvFile = new Blob([csv], {type: "text/csv"});

  // Download link
  downloadLink = document.createElement("a");

  // File name
  downloadLink.download = filename;

  // Create a link to the file
  downloadLink.href = window.URL.createObjectURL(csvFile);

  // Hide download link
  downloadLink.style.display = "none";

  // Add the link to DOM
  document.body.appendChild(downloadLink);

  // Click download link
  downloadLink.click();
}

exportToCsv = () => {
  var filename = 'grid.csv'
  var csv = [];
  const spreadsheetData = this.getData();

  for (var i = 1; i < spreadsheetData.length; i++) {
    var row = []
    var col = spreadsheetData[i];
    
    for (var j = 1; j < col.length; j++) 
        row.push(col[j]);
    
    csv.push(row.join(","));     
  }
  // Download CSV file
  downloadCSV(csv.join("\n"), filename);
}

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
        addRow(parseInt(indices[1]), "bottom");  
    }
  });

  addColumnBtn.addEventListener("click", function(e) {
    if (e.target) {
        const indices = index.split("-");
        addColumn(parseInt(indices[2]), "right");  
    }
  });

  deleteRowBtn.addEventListener("click", function(e) {
    if (e.target) {
        const indices = index.split("-");
        deleteRow(parseInt(indices[1]));  
    }
  });

  deleteColumnBtn.addEventListener("click", function(e) {
    if (e.target) {
        const indices = index.split("-");
        deleteColumn(parseInt(indices[2]));  
    }
  });

  importCsvBtn.addEventListener("click", function(e) {
    importFromCsv();
  });

  exportCsvBtn.addEventListener("click", function() {
    exportToCsv();
  });
  
}

calculateExp = formula => {
  let formulaArr = [];
  let input = [];
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
  }
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
  // process until we are done
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

createCellObservable = cellId => {
  return rxjs.fromEvent(document.getElementById(cellId), "focusout");
};

createObservables = formula => {
  if(formula.startsWith("=SUM")){
    formula = formula.replace("SUM(","");
    formula = formula.replace(")","");
  }
  let formulaArr = formula.split(/[=*/%+-]+/);
  let currentObs = [];
  //["A1", "A2"] r-1-1, r-2-1
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

  // attach focusout event listener to whole table body container
  tableBody.addEventListener("focusout", function(e) {
    if (e.target && e.target.nodeName === "TD") {
      let item = e.target;
      const indices = item.id.split("-");
      let data = getData();
      let currentCellData = item.innerHTML;
      data[indices[1]][indices[2]].actualValue = currentCellData;
      if(currentCellData.startsWith("=")){
        let calcValue = calculateExp(currentCellData);
        data[indices[1]][indices[2]].actualValue = item.innerHTML;
        data[indices[1]][indices[2]].displayValue = calcValue;
        document.getElementById(item.id).innerHTML = calcValue;
        createObservables(
            data[indices[1]][indices[2]].actualValue
        ).forEach(observable => {
          observable.subscribe(() => {
            document.getElementById(item.id).innerHTML = calculateExp(
                data[indices[1]][indices[2]].actualValue
            );
            data[indices[1]][indices[2]].displayValue = document.getElementById(item.id).innerHTML;
          });
        });
      }else{
        data[indices[1]][indices[2]].actualValue = item.innerHTML;
        data[indices[1]][indices[2]].displayValue = item.innerHTML;
        saveData(data);
      }
    }
  });

  // Attach click event listener to table body
  tableBody.addEventListener("click", function(e) {
    if (e.target) {
      index = e.target.id;
      if (e.target.className === "dropbtn") {
        const idArr = e.target.id.split("-");
        document
          .getElementById(`row-dropdown-${idArr[2]}`)
          .classList.toggle("show");
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

  // Attach click event listener to table headers
  tableHeaders.addEventListener("click", function(e) {
    if (e.target) {
      if (e.target.className === "column-header-span") {
        sortColumn(parseInt(e.target.parentNode.id.split("-")[2]));
      }
      if (e.target.className === "dropbtn") {
        const idArr = e.target.id.split("-");
        document
          .getElementById(`col-dropdown-${idArr[2]}`)
          .classList.toggle("show");
      }
      if (e.target.className === "col-insert-left") {
        const indices = e.target.parentNode.id.split("-");
        addColumn(parseInt(indices[2]), "left");
      }
      if (e.target.className === "col-insert-right") {
        const indices = e.target.parentNode.id.split("-");
        addColumn(parseInt(indices[2]), "right");
      }
      if (e.target.className === "col-delete") {
        const indices = e.target.parentNode.id.split("-");
        deleteColumn(parseInt(indices[2]));
      }
    }
  });
};
createSpreadsheet();
createControls();

// Close the dropdown menu if the user clicks outside of it
window.onclick = function(event) {
  if (!event.target.matches(".dropbtn")) {
    var dropdowns = document.getElementsByClassName("dropdown-content");
    var i;
    for (i = 0; i < dropdowns.length; i++) {
      var openDropdown = dropdowns[i];
      if (openDropdown.classList.contains("show")) {
        openDropdown.classList.remove("show");
      }
    }
  }
};
window.onload = () => {
}

