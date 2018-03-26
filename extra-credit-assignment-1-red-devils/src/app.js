var table = document.getElementById("spreadsheetTable"); //We are getting the table using the ID Attribute and storing it in a table


for (let i = 0; i < 11; i++) {
    var row = document.getElementById("spreadsheetTable").insertRow(-1); //JS Function to insert row
    for (let j = 0; j < 11; j++) {
        if (j < 27) {
            var letter = String.fromCharCode("A".charCodeAt(0) + j - 1); //Getting the Alphabet Numbers
        } else {
            var letter = "A" + String.fromCharCode("A".charCodeAt(0) + j - 27);
        }

        row.insertCell(-1).innerHTML = i && j ? "<input id='" + letter + i + "'  class='tableCells'/>" : i || letter; // inserting a new row
    }
};

//FUNCTION TO INSERT NEW ROWS
var insertButton = document.getElementById("insert"); //Get insert button and add Event listener to it

var insertRows = function() {
    var rows = document.getElementById("spreadsheetTable").getElementsByTagName("tr").length; //Get no.of rows
    var columns = document.getElementById('spreadsheetTable').rows[0].cells.length; //Get no.of Columns
    var newRow = document.getElementById("spreadsheetTable").insertRow(-1); //Insert New Column
    for (let k = 0; k < columns; k++) {
        if (k < 27) {
            var letter = String.fromCharCode("A".charCodeAt(0) + k - 1);
        } else {
            var letter = "A" + String.fromCharCode("A".charCodeAt(0) + k - 27);
        }
        //Conditional Ternary operator to add cells/letters/numbers
        newRow.insertCell(-1).innerHTML = k ? "<input id='" + letter + rows + "' class='tableCells'/>" : rows;
    }

};
insertButton.addEventListener('click', insertRows);

//This function will help us insert a new row into a selected position
let insertatButton = document.getElementById("insertat"); //Adding an event listener to the button


var insertRowsAt = function() {
    let rowLength = document.getElementById("spreadsheetTable").getElementsByTagName("tr").length; //The line helps us get the number of rows
    let columns = document.getElementById('spreadsheetTable').rows[0].cells.length; //The line helps us get the number of columns
    let rowNo = document.getElementById("insertRowNo").value; //The row number at which a new row needs to be inserted
    let newRow = document.getElementById("spreadsheetTable").insertRow(rowNo); //Inserting a new Column


    for (let k = 0; k < columns; k++) {
        if (k < 27) {
            var letter = String.fromCharCode("A".charCodeAt(0) + k - 1);
        } else {
            var letter = "A" + String.fromCharCode("A".charCodeAt(0) + k - 27);
        }
        newRow.insertCell(-1).innerHTML = k ? "<input id='" + letter + rowNo + "' class='tableCells'/>" : rowNo;
    }
    //Updating the cells and the row nos
    for (let i = rowNo; i < rowLength + 1; i++) {

        for (let k = 0; k < columns; k++) {
            if (k < 27) {
                var letter = String.fromCharCode("A".charCodeAt(0) + k - 1);
            } else {
                var letter = "A" + String.fromCharCode("A".charCodeAt(0) + k - 27);
            }
            //Ternary operator to add letters/numbers
            table.rows[i].cells[k].innerHTML = k ? "<input id='" + letter + i + "' value='" + "'/>" : i;
        }
    }
};
insertatButton.addEventListener('click', insertRowsAt); // after the button is clicked the insertRowsAt function is called


//Function to insert at the very last position
let insertColumnButton = document.getElementById("insertColumn"); //Getting and adding an event listener

let insertColumn = function() {
    let rows = document.getElementById("spreadsheetTable").getElementsByTagName("tr").length; //No of rows
    let columns = document.getElementById('spreadsheetTable').rows[0].cells.length; //No of columns

    for (let x = 0; x < rows; x++) {
        if (columns < 27) {
            var letter = String.fromCharCode("A".charCodeAt(0) + columns - 1);
        } else {
            var letter = "A" + String.fromCharCode("A".charCodeAt(0) + columns - 27);
        }
        //Ternary operator
        table.rows[x].insertCell(-1).innerHTML = x ? "<input id='" + letter + x + "' class='tableCells'/>" : letter;
    }
};
insertColumnButton.addEventListener('click', insertColumn);

//To delete the last row
let deleteRowButton = document.getElementById("deleteRow"); //Adding an Event Listener

var deleteRow = function() {
    let rowLength = document.getElementById("spreadsheetTable").getElementsByTagName("tr").length; //Get no.of rows
    table.deleteRow(rowLength - 1); //Deleting using deleterow function
};
deleteRowButton.addEventListener('click', deleteRow);

//To delete the last column
let deleteColumnButton = document.getElementById("deleteColumn");

var deleteColumn = function() {
    let lastRow = document.getElementById("spreadsheetTable").getElementsByTagName("tr"); //The row count
    let columnLength = document.getElementById('spreadsheetTable').rows[0].cells.length; //The column Count
    for (let y = 0; y < lastRow.length; y++) {
        table.rows[y].deleteCell(columnLength - 1); //Deleting the last cell of each row
    }
};
deleteColumnButton.addEventListener('click', deleteColumn);
//Adding events to the cells of the table
var allCells = document.getElementsByClassName("tableCells"); //Getting using class names 

for (let z = 0; z < allCells.length; z++) {
    //Adding onblur event
    allCells[z].onblur = function(e) {
        calculateAll(); //Evaluates on each cell
    };
};

//Evaluating excel functions on all cells 
var calculateAll = function() {
    for (var z = 0; z < allCells.length; z++) {
        //Get the cells that have strings starting with '=' or has attribute 'formula' which we are assigning 
        if (allCells[z].value.charAt(0) === "=" || allCells[z].hasAttribute('formula')) {
            //Store the string without '=' by suing substring
            if (allCells[z].value.charAt(0) === "=") {
                var formula = allCells[z].value.substring(1);
            } else {
                var formula = allCells[z].getAttribute('formula');
            }

            //Getting the values at the given cells using regex
            var expression = formula.replace(/([A-Z]+[0-9]+)/g, "parseFloat(document.getElementById('$1').value)");
            //Evaluate the value
            allCells[z].value = isNaN(parseFloat(eval(expression))) ? "" : parseFloat(eval(expression));
            //Set an attribute
            allCells[z].setAttribute('formula', formula);
            allCells[z].onfocus = function(e) {
                //storing the formula onfocus
                e.target.value = "=" + formula;
            }
        } else { allCells[z].value = isNaN(parseFloat(allCells[z].value)) ? allCells[z].value : parseFloat(allCells[z].value); }
    };
};

//for the CSV File
function download_csv(csv, filename) {
    var csvFile;
    var downloadLink;

    // CSV FILE
    csvFile = new Blob([csv], { type: "text/csv" });

    // Download link
    downloadLink = document.createElement("a");

    // The name of the file
    downloadLink.download = filename;

    // We have to create a link to the file
    downloadLink.href = window.URL.createObjectURL(csvFile);

    //To make sure that the link is not displayed
    downloadLink.style.display = "none";

    // Add the link to your DOM
    document.body.appendChild(downloadLink);


    downloadLink.click();
}

function export_table_to_csv(html, filename) { //exporting the table to csv
    var csv = [];
    var rows = document.querySelectorAll("table tr");

    for (var i = 1; i < rows.length; i++) {
        var row = [],
            cols = rows[i].querySelectorAll(".tableCells");

        for (var j = 0; j < cols.length; j++)
            row.push(cols[j].value);

        csv.push(row.join(","));
    }

    // Downloading the CSV
    download_csv(csv.join("\n"), filename);
}

document.getElementById("downloadButton").addEventListener("click", function() { // adding an event listener to the download button on clicking will call the anonymous function
    var html = document.querySelector("table").outerHTML;
    export_table_to_csv(html, "table.csv");
});