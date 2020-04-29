//Initialising the number of rows and columns
var init_rows = 10;
var init_columns = 6;

//Fixing the max number of columns allowed in the sheet
var max_columns = 26;

// Created the subjects for emiting the values entered in the cells
const sub = new rxjs.Subject();
const rowObs = new rxjs.Subject();

// This function highlights the row selected and vice versa 
const selectedRow = new Set();
const selectRow = (x) => {
    if (x.classList.contains("highlight")) {
        x.classList.remove("highlight");
        selectedRow.delete(x);
    }
    else {
        x.classList.add("highlight");
        selectedRow.add(x);
    }
}

// This function highlights the column selected and vice versa 
const selectedColumn = new Set();
const selectCol = (x) => {
    if (x.classList.contains("highlight")) {
        selectedColumn.delete(x.id);
    }
    else {
        selectedColumn.add(x.id);
    }
    //x.classList.toggle("highlight");
    for (let i = 0; i < init_rows; i++) {
        let id = x.id[0];
        let newid = id + i;
        // console.log(newid);
        let col = document.getElementById(newid);
        col.classList.toggle("highlight");
    }
}

// This function helps to check the formula entered and assign the operator to carry out the functionality user entered
const calculate = (td) => {
    //rxjs event has been started to monitor the value entered in the cell
    rxjs.fromEvent(td, 'input').pipe(rxjs.operators.debounceTime(800)).subscribe(x => {
        // check statement for the formula starts with =SUM()
        if (td.innerText.startsWith("=SUM(") && td.innerText.endsWith(")")) {
            let actualStr = td.innerText.substring(5, td.innerText.length - 1);
            // let actualStr = initalStr.substring(1, initalStr.length - 1);
            let arr = [];
            actualStr.split(":").forEach(x => {
                if (x.length > 1)
                    arr.push(x);
            });
            if (arr.length == 2) {
                sum(td, arr);
            }
        }
        // check statement for the formula starts with =
        else if (td.innerText.startsWith("=", 0) && td.innerText.length > 5 && !td.innerText.includes("=SUM")) {
            let initString = td.innerText.substring(1, td.innerText.length);
            let signs = ["+", "-", "*", "/"];
            let operator = "";
            // assigns the operator written in the formula entered
            for (let i = 0; i < signs.length; i++) {
                if (initString.includes(signs[i])) {
                    operator = signs[i];
                }
            }
            let arra = [];
            initString.split(operator).forEach(x => {
                arra.push(x);
            })
            if (arra.length == 2) {
                switch (true) {
                    case (operator === "+"): {
                        td.setAttribute("isFormula", "true")
                        td.setAttribute("type", "sum")
                        operate(td, "+")
                        break;
                    }
                    case (operator === "-"): {
                        td.setAttribute("isFormula", "true")
                        td.setAttribute("type", "diff")
                        operate(td, "-")
                        break;
                    }
                    case (operator === "*"): {
                        td.setAttribute("isFormula", "true")
                        td.setAttribute("type", "mul")
                        operate(td, "*")
                        break;
                    }
                    case (operator === "/"): {
                        td.setAttribute("isFormula", "true")
                        td.setAttribute("type", "div")
                        operate(td, "/")
                        break;
                    }
                    default:
                        {
                            window.alert("Not a valid formula");
                        }
                }
            }
        }
        // check if the formula has been deleted and to unsubscribe it
        else if (x.inputType == "deleteContentBackward" && td.getAttribute("isFormula") == "true") {
            td.removeAttribute("isFormula");
            td.removeAttribute("type");
        }
        sub.next(x.target);
    });
}

// This function operates the functionality entered in the cell
const operate = (td, type) => {
    let initString = td.innerText.substring(1, td.innerText.length);
    let signs = ["+", "-", "*", "/"];
    let operator = "";
    for (let i = 0; i < signs.length; i++) {
        if (initString.includes(signs[i])) {
            operator = signs[i];
        }
    }
    let arra = [];
    initString.split(operator).forEach(x => {
        arra.push(x);
    })
    let a = document.getElementById(arra[0]);
    let b = document.getElementById(arra[1]);
    // Executes the formula entered in the cell through the below mentioned functionalities
    let observer = sub.subscribe(x => {
        let sum = 0;
        switch (true) {
            case (td.getAttribute("isFormula") && td.getAttribute("type") == "sum"): {
                td.innerText = parseInt(a.innerText) + parseInt(b.innerText);
                break;
            }
            case (td.getAttribute("isFormula") && td.getAttribute("type") == "diff"): {
                td.innerText = parseInt(a.innerText) - parseInt(b.innerText);
                break;
            }
            case (td.getAttribute("isFormula") && td.getAttribute("type") == "mul"): {
                td.innerText = parseInt(a.innerText) * parseInt(b.innerText);
                break;
            }
            case (td.getAttribute("isFormula") && td.getAttribute("type") == "div"): {
                td.innerText = parseInt(a.innerText) / parseInt(b.innerText);
                break;
            }
            default:
                {
                    // Unsubscription of the subject
                    observer.unsubscribe();
                }
        }
    });
}

// This function is used to observe and carry out the consecutive sum through the rows as well columns in the specified cell range
const sum = (td, arr) => {
    if (arr[0].charAt(0) == arr[1].charAt(0)) {
        let col = arr[0].charAt(0);
        td.setAttribute("isFormula", "true")
        let start = parseInt(arr[0].substring(1, arr[0].length));
        let end = parseInt(arr[1].substring(1, arr[1].length));
        let rowObserver = rowObs.subscribe(x => {
            if (x < end && x >= start) {
                end = parseInt(end) + 1;
            } else if (x < start) {
                start = parseInt(start) + 1;
                end = parseInt(end) + 1;
            }
        });
        let observer = sub.subscribe(x => {
            if (td.getAttribute("isFormula")) {
                let sum = 0;
                for (i = start; i <= end; i++) {
                    sum = sum + parseInt(document.getElementById(col + i).innerText);
                }
                td.innerText = sum;
            } else {
                observer.unsubscribe();
                rowObserver.unsubscribe();
            }
        });
    } else if (arr[0].substring(1, arr[0].length) == arr[1].substring(1, arr[1].length)) {
        td.setAttribute("isFormula", "true")
        let observer = sub.subscribe(x => {
            if (td.getAttribute("isFormula")) {
                let sum = 0;
                let start = arr[0].charCodeAt(0);
                let end = arr[1].charCodeAt(0);
                let val = arr[0].substring(1, arr[0].length);
                for (i = start; i <= end; i++) {
                    sum = sum + parseInt(document.getElementById(String.fromCharCode(i) + val).innerText);
                }
                td.innerText = sum;
            } else { observer.unsubscribe() }
        });
    } else {
        console.log("invalid");
    }
}

// A function to load constant rows & columns as initiated 
const reload = () => {
    let body = document.getElementsByTagName("body")[0];
    let table = document.createElement("table");

    for (let i = 0; i < init_rows; i++) {
        let tr = document.createElement("tr");
        tr.setAttribute("id", i);
        for (let j = 0; j < init_columns; j++) {
            let td = document.createElement("td");
            if (i != 0) {
                td.setAttribute("id", String.fromCharCode(j + 64) + i);
            }
            else {
                td.setAttribute("id", String.fromCharCode(j + 64) + i);
            }
            if (i == 0 & j > 0) {
                let text = document.createTextNode(String.fromCharCode(j + 64));
                td.addEventListener("click", function () {
                    selectCol(td);
                }, false);
                td.appendChild(text);
            }
            else if (j == 0 & i > 0) {
                let text = document.createTextNode(i);
                td.addEventListener("click", function () {
                    selectRow(tr);
                }, false);
                td.appendChild(text);
            }
            else if (i == 0 & j == 0) {
                td.setAttribute("contenteditable", "false");
            }
            else {
                td.setAttribute("contenteditable", "true");
                calculate(td);
            }
            tr.appendChild(td);
        }
        table.appendChild(tr);
    }
    body.appendChild(table);
}

// Calls the function when window loads
window.onload = reload();

//Function add row to add the rows below the selected row and also alerts the user on selection of row.
// Also the observer emits the values entered in the cells of the new row added
const addRow = () => {

    if (selectedRow.size != 1 || selectedRow.size == 0) {
        alert("Please select one row");
    }
    else if (selectedRow.size === 1) {
        let table = document.getElementsByTagName("table")[0];
        let iterator = selectedRow.values();
        let row = iterator.next().value;
        let index = row.rowIndex + 1;
        rowObs.next(index -1);
        let newRow = table.insertRow(index);
        newRow.setAttribute("id", index);
        console.log(index);
        for (let i = 0; i < init_columns; i++) {
            let newcell = newRow.insertCell(i);
            if (i == 0) {
                newcell.setAttribute("contenteditable", "false");
                newcell.setAttribute("id", index);
                newcell.addEventListener("click", function () {
                    selectRow(newRow);
                }, false);
            }
            else {
                newcell.setAttribute("contenteditable", "true");
                newcell.setAttribute("id", String.fromCharCode(i + 64) + index);
                selectedColumn.forEach(x=>{
                    //let sel = x.id;
                    let newid = x.charCodeAt(0) - 64;
                    if(newid == i)
                    newcell.setAttribute("class","highlight");
                });
                calculate(newcell);
            }
        }
        init_rows = init_rows + 1;
        rearrangeTableAdd(index);
        // newRow.appendChild(table);
    }
    console.log(init_rows);
}

// Function to reaarange the table when the new row has been added 
const rearrangeTableAdd = (index) => {
    let table = document.getElementsByTagName("table")[0];
    //let row = index;
    let x = table.rows;
    for (let i = index; i < init_rows; i++) {
        for (let j = 0; j < init_columns; j++) {
            let y = x[i].cells;
            y[0].innerText = i;
            y[j].setAttribute("id", String.fromCharCode(j + 64) + i);
        }
    }
}

// Function to delete the row selected based on the selected index and also alerts the user on deletion of the last few rows
const deleteRow = () => {
    if (init_rows < 3) {
        alert("You are not allowed to delete the last few rows");
    }
    else {
        if (selectedRow.size != 1 || selectedRow.size == 0) {
            alert("Please select one row to delete");
        }
        else {
            let table = document.getElementsByTagName("table")[0];
            let iterator = selectedRow.values();
            let row = iterator.next().value;
            let index = row.rowIndex;
            table.deleteRow(index);
            selectedRow.clear();
            init_rows = init_rows - 1;
            rearrangeTableDelete(index);
        }
    }
    console.log(init_rows);
}

// function to reaarange the table when a row is deleted 
const rearrangeTableDelete = (index) => {
    let table = document.getElementsByTagName("table")[0];
    //let row = index;
    let x = table.rows;
    for (let i = index; i < init_rows; i++) {
        for (let j = 0; j < init_columns; j++) {
            let y = x[i].cells;
            y[0].innerText = i;
            y[j].setAttribute("id", String.fromCharCode(j + 64) + i);
        }
    }
}

//Event Listener to add/Delete row by getting the element by ID
document.getElementById("addRow").addEventListener("click", addRow, false);
document.getElementById("deleteRow").addEventListener("click", deleteRow, false);

//Function add column to add the column next to the selected column and also alerts the user on selection of column.
const addColumn = () => {
    if (init_columns > max_columns) {
        alert("You are not allowed to add any more new columns");
    }
    else {
        if (selectedColumn.size != 1 || selectedColumn.size == 0) {
            alert("Please select one column");
        }
        else {
            let table = document.getElementsByTagName("table")[0];
            let iterator = selectedColumn.values();
            let col = iterator.next().value;
            let newid = col.charCodeAt(0) - 64;
            let rel = newid + 1;
            for (let i = 0; i < init_rows; i++) {
                //let id = column.id(0);
                let id = col[0];
                let column = id + i;
                let oldtd = document.getElementById(column);
                let td = document.createElement("td");

                if (i == 0) {
                    td.setAttribute("contenteditable", "false");
                    td.setAttribute("id", String.fromCharCode(rel + 64) + i);
                    td.addEventListener("click", function () {
                        selectCol(td);
                    }, false);
                }
                else {
                    td.setAttribute("contenteditable", "true");
                    td.setAttribute("id", String.fromCharCode(rel + 64) + i);
                    calculate(td);
                }

                oldtd.insertAdjacentElement('afterend', td);
            }
            init_columns = init_columns + 1;
            rearrangeTableColumn(rel);
            console.log(init_rows);
        }
    }
}

// Function to rearrange the table once the column has been added
const rearrangeTableColumn = (rel) => {
    let table = document.getElementsByTagName("table")[0];
    //let row = index;
    let x = table.rows;
    for (let i = 0; i < init_rows; i++) {
        for (let j = rel; j < init_columns; j++) {
            let y = x[i].cells;
            if (i == 0)
                y[j].innerText = String.fromCharCode(j + 64);
            y[j].setAttribute("id", String.fromCharCode(j + 64) + i);
            // if (selectedCol.has(String.fromCharCode(j + 64))) {
            //     y[j].classList.add("highlight");
        }
    }
}

// Function to delete the selected column and is iterated in all the rows. Also it alerts
// the user on deletion of the last few columns
const deleteColumn = () => {
    if (init_columns < 3) {
        alert("You are not allowed to delete the last few columns");
    }
    else {
        if (selectedColumn.size != 1 || selectedColumn.size == 0) {
            alert("Please select one column to delete");
        }
        else {
            let table = document.getElementsByTagName("table")[0];
            //let row = index;
            let iterator = selectedColumn.values();
            let col = iterator.next().value;
            let newid = col.charCodeAt(0) - 64;
            let delcol = newid;
            let x = table.rows;
            for (let i = 0; i < init_rows; i++) {
                x[i].deleteCell(delcol);
            }
            selectedColumn.clear();
            init_columns = init_columns - 1;
            rearrangeTableDeleteCol(delcol);
        }
    }
    console.log(init_rows);
}

// Function to reaarange the column once the column has been deleted
const rearrangeTableDeleteCol = (delcol) => {
    let table = document.getElementsByTagName("table")[0];
    //let row = index;
    let x = table.rows;
    for (let i = 0; i < init_rows; i++) {
        for (let j = delcol; j < init_columns; j++) {
            let y = x[i].cells;
            if (i == 0)
                y[j].innerText = String.fromCharCode(j + 64);
            y[j].setAttribute("id", String.fromCharCode(j + 64) + i);
        }
    }
}

//Event Listener to add/Delete Column by getting the element by ID
document.getElementById("addColumn").addEventListener("click", addColumn, false);
document.getElementById("deleteColumn").addEventListener("click", deleteColumn, false);

//This function will export the table data into CSV
const export_table_to_csv = (html, filename) => {
    let csv = [];
    let rows = document.querySelectorAll("table tr");

    //The below for loop loops through every row to traverse through each td (basically cell)
    // for (let i = 1; i < rows.length; i++) { -- if header need not be exported
    for (let i = 0; i < rows.length; i++) {
        let row = [], cols = rows[i].querySelectorAll("td");
        //This loop traverses through each input field of the table
        //for (let j = 1; j < cols.length; j++) { -- if header need not be exported
        for (let j = 0; j < cols.length; j++) {
            let newid = cols[j].id;
            let value = document.getElementById(newid).innerText;
            row.push(value);
        }
        csv.push(row.join(","));
    }
    //This will call the download CSV function
    download_csv(csv.join("\n"), filename);
}

const exportcsv = () => {
    let html = document.querySelector("table").outerHTML;
    export_table_to_csv(html, "table.csv");
}

//this function will download the CSV upon clicking the download CSV button
const download_csv = (csv, filename) => {
    let csvFile;
    let downloadLink;

    //This is basically the csv file
    csvFile = new Blob([csv], { type: "text/csv" });

    //This is the download link
    downloadLink = document.createElement("a");

    //This is the filename that will of the CSV that will be downloaded
    downloadLink.download = filename;

    //A link to the file needs to be create and the link should not be displayed
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = "none";

    //The link needs to be appended to the DOM
    document.body.appendChild(downloadLink);
    downloadLink.click();
}

// event listener to export the current table as csv
document.getElementById("export").addEventListener("click", exportcsv, false);

//method for pop up of the upload csv dialog box
const div_show = () => {
    document.getElementById('upload').style.display = "block";
}

// event listener to show the pop up of the upload csv dialog box 
document.getElementById("uploadpopup").addEventListener("click", div_show, false);

// method to hide the pop up of the upload csv dialog box
const div_hide = () => {
    document.getElementById('upload').style.display = "none";
}

// event listener to hide the pop up of the upload csv dialog box
document.getElementById("uploadhide").addEventListener("click", div_hide, false);


// Function to import the csv file and append it in our spreadsheet.
const Upload = () => {
    let fileUpload = document.getElementById("fileUpload");
    let regex = /^([a-zA-Z0-9\s_\\.\-:])+(.csv|.txt)$/;
    if (regex.test(fileUpload.value.toLowerCase())) {
        if (typeof (FileReader) != "undefined") {
            let reader = new FileReader();
            reader.onload = function (e) {
            let table = document.getElementsByTagName("table")[0];
                let x = table.rows;
                for (j = 0; j < x.length; j++) {
                    // let index = j + 1;
                    table.deleteRow(j);
                    j = j - 1;
                }
                let body = document.getElementsByTagName("body")[0];
                selectedColumn.clear();
                selectedRow.clear();
                body.removeChild(table);
                let rows = e.target.result.split("\n");
                init_rows = rows.length;
                for (let i = 0; i < rows.length; i++) {
                    let cell = rows[0].split(",");
                    if (cell.length > max_columns) {
                        alert("Showing maximum possible columns");
                        init_columns = max_columns + 1;
                        break;
                    }
                    else {
                        let temp = cell.length + 1;
                        init_columns = temp;
                    }
                }
                reload();
                let table1 = document.getElementsByTagName("table")[0];
                let x1 = table1.rows;
                let r = e.target.result.split("\n");
                //init_rows = rows.length;
                for (let l = 0, i = 1; l < r.length, i < init_rows; i++ , l++) {
                    let cell = r[l].split(",");
                    if (cell.length < max_columns) {
                        for (let k = 0, j = 1; k < cell.length, j < init_columns; k++ , j++) {
                            let y = x1[i].cells;
                            y[j].innerText = cell[k];
                        }
                    }
                    else {
                        for (let k = 0, j = 1; k < max_columns, j < init_columns; k++ , j++) {
                            let y = x1[i].cells;
                            y[j].innerText = cell[k];
                        }
                    }
                }
            }
            reader.readAsText(fileUpload.files[0]);
            div_hide();

        } else {
            alert("This browser does not support HTML5.");
            div_hide();
        }
    } else {
        alert("Please upload a valid CSV file.");
        div_hide();
    }
}

// event listener to upload the csv from the local file system
document.getElementById("uploadcsv").addEventListener("click", Upload, false);