import React from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

import logo from "./logo.svg";
import "./App.css";

function App() {
  const makeExcel = () => {
    // https://github.com/exceljs/exceljs#interface
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("ExcelJS sheet");
    const row = worksheet.getRow(5);
    row.getCell(1).value = 5; // A5's value set to 5
    row.getCell("C").value = new Date(); // C5's value set to now

    workbook.xlsx.writeBuffer().then((buffer) => {
      saveAs(
        new Blob([buffer], { type: "application/octet-stream" }),
        `cecdata.xlsx`
      );
    });
  };
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <button onClick={makeExcel}>Export Excel</button>
      </header>
    </div>
  );
}

export default App;
