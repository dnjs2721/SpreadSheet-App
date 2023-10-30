const spreadSheetContainer = document.querySelector("#spreadsheet-container");
const exportBtn = document.querySelector("#export-btn"); 
const ROWS = 10;
const COLUMNS = 10;
const spreadsheet = [];

class Cell {
    constructor(isHeader, disabled, data, row, column, rowName, columnName, active = false) {
        this.isHeader = isHeader;
        this.disabled = disabled;
        this.data = data;
        this.row = row;
        this.column = column;
        this.rowName = rowName;
        this.columnName = columnName;
        this.active = active;
    }
}

exportBtn.onclick = function() {
    let csv = "";
    for (let i = 0; i < spreadsheet.length; i++) {
        if (i === 0) continue;
        csv += 
            spreadsheet[i]
                .filter((item) => !item.isHeader)
                .map((item) => item.data)
                .join(",") + "\r\n";
    }
    const csvObj = new Blob([csv]);
    const csvUrl = URL.createObjectURL(csvObj);
    console.log(csvUrl);

    const a = document.createElement("a");
    a.href = csvUrl;
    a.download = "Spreadsheet File Name.excel";
    a.click();
}

initSpreadsheet();

function initSpreadsheet() {
    for (let i = 0; i < ROWS; i++) {
        let spreadsheetRow = [];
        for (let j = 0; j < COLUMNS; j++) {
            let cellData = '';
            let header = false;
            let disabled = false;

            const rowName = i;
            const columnName = String.fromCharCode(64 + j);

            if (i === 0 && j === 0) {
                header = true;
                disabled = true;
            }

            if (j === 0 && i >= 1) {
                cellData = rowName;
                header = true;
                disabled = true;
            }

            if (i === 0 && j >= 1) {
                cellData = columnName;
                header = true;
                disabled = true;
            } 

    
            const cell = new Cell(header, disabled, cellData, i, j, rowName, columnName,false);
            spreadsheetRow.push(cell);
        }
        spreadsheet.push(spreadsheetRow);
    }
    drawSheet();
}

function createCellEl(cell) {
    const cellEl = document.createElement("input");
    cellEl.className = "cell";
    cellEl.id = "cell_" + cell.row + cell.column;
    cellEl.value = cell.data;
    cellEl.disabled = cell.disabled;

    if (cell.isHeader) cellEl.classList.add("header");

    cellEl.addEventListener("focus", () => handelCellFocus(cell));
    cellEl.onchange = (e) => handleOnChange(e.target.value, cell);

    return cellEl;
}

function handleOnChange(value, cell) {
    cell.data = value;
}

function handelCellFocus(cell) {
    clearHeaderActiveStatus();
    const columnHeader = spreadsheet[0][cell.column];
    const rowHeader = spreadsheet[cell.row][0];

    const columnHeaderEl = getElFromRowCol(columnHeader.row, columnHeader.column);
    const rowHeaderEl = getElFromRowCol(rowHeader.row, rowHeader.column);

    columnHeaderEl.classList.add("active");
    rowHeaderEl.classList.add("active");
    document.querySelector("#cell-status").innerHTML = cell.columnName + "" + cell.rowName;
}

function clearHeaderActiveStatus() { 
    const headers = document.querySelectorAll('.header.active');
    headers.forEach((header) => {
        header.classList.remove('active');
    })
}

function getElFromRowCol(row, col) {
    return document.getElementById("cell_" + row + col);
}

function drawSheet() {
    for (let i = 0; i < spreadsheet.length; i++) {
        const rowContainerEl = document.createElement("div");
        rowContainerEl.className = "cell-row";
        for (let j = 0; j < spreadsheet[i].length; j++) {
            const cell = spreadsheet[i][j];
            const cellEl = createCellEl(cell);
            rowContainerEl.append(cellEl);
        }
        spreadSheetContainer.append(rowContainerEl);
    }
}