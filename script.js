let table = document.getElementsByClassName("sheet-body")[0],
  rows = document.getElementsByClassName("rows")[0],
  columns = document.getElementsByClassName("columns")[0];
tableExists = false;

const generateTable = () => {
  let rowsNumber = parseInt(rows.value),
    columnsNumber = parseInt(columns.value);

  //todo Displaying the Alert Message for the user if the input fields are empty
  if (!rowsNumber && !columnsNumber) {
    Swal.fire({
      icon: "error",
      title: "Oops...",
      text: "You Must Enter the Rows and Columns Number !!",
    });
  } else if (!rowsNumber && columnsNumber) {
    Swal.fire({
      icon: "error",
      title: "Oops...",
      text: "You Must Enter the Rows Number !!",
    });
  } else if (!columnsNumber && rowsNumber) {
    Swal.fire({
      icon: "error",
      title: "Oops...",
      text: "You Must Enter the Columns Number !!",
    });
  }

  table.innerHTML = "";
  for (let i = 0; i < rowsNumber; i++) {
    var tableRow = "";
    for (let j = 0; j < columnsNumber; j++) {
      tableRow += `<td contenteditable></td>`;
    }
    table.innerHTML += tableRow;

    //* Displaying Success Message
    Swal.fire({
      position: "center",
      icon: "success",
      title: "Your work has been saved",
      timer: 1500,
      showConfirmButton: true,
    });
  }
  if (rowsNumber > 0 && columnsNumber > 0) {
    tableExists = true;
  }
};

const ExportToExcel = (type, fn, dl) => {
  if (!tableExists) {
    //todo Displaying the Alert Message for the user if he didn't generate the table
    Swal.fire({
      icon: "error",
      title: "Oops...",
      text: "You Must Generate the Table First Then Export !!",
    });
    return;
  }
  var elt = table;
  var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
  return dl
    ? XLSX.write(wb, { bookType: type, bookSST: true, type: "base64" })
    : XLSX.writeFile(wb, fn || "MyNewSheet." + (type || "xlsx"));
};
