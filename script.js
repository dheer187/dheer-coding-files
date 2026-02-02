const processButton = document.getElementById("processButton");
const excelFileInput = document.getElementById("excelFile");
const textBox = document.getElementById("textBox");
const downloadLink1 = document.getElementById("downloadExcel1");
const downloadLink2 = document.getElementById("downloadExcel2");

function createExcel(data, fileName) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const downloadLink = fileName === "output1.xlsx" ? downloadLink1 : downloadLink2;
  downloadLink.href = url;
  downloadLink.style.display = "block";
  downloadLink.textContent = `Download ${fileName}`;
}

function processFiles() {
  const excelFile = excelFileInput.files[0];
  const textData = textBox.value;

  if (!excelFile || !textData) {
    alert("Please upload a file and enter text data.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Create two new Excel files
    const newData1 = [...jsonData, ["Text Data", textData]]; // Add text data at the bottom of file
    const newData2 = jsonData.map(row => [...row, "Pasted Data"]); // Copy data and add a new column

    createExcel(newData1, "output1.xlsx");
    createExcel(newData2, "output2.xlsx");
    alert("Files are ready for download!");
  };
  reader.readAsArrayBuffer(excelFile);
}

processButton.addEventListener("click", processFiles);