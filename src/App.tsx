import React, { useRef } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { LineChart, Line } from "recharts";

import "./App.css";

function App() {
  const chartRef = useRef<any>(null);

  const renderLineChart = (
    <LineChart ref={chartRef} width={400} height={400} data={data}>
      <Line type="monotone" dataKey="uv" stroke="#8884d8" />
    </LineChart>
  );

  const makeExcel = async () => {
    // https://github.com/exceljs/exceljs#interface
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("ExcelJS sheet");
    const row = worksheet.getRow(5);
    row.getCell(1).value = 5; // A5's value set to 5
    row.getCell("C").value = new Date(); // C5's value set to now

    // turn the chart into an image and embed it
    if (chartRef) {
      // we should always have the chart but if it's null for some reason skip it instead of breaking
      const chartSvg = chartRef.current.container.children[0];

      const pngData = await svgToPng(chartSvg, 500, 500);

      // add image to workbook by base64
      const chartImageId2 = workbook.addImage({
        base64: pngData,
        extension: "png",
      });

      // insert an image over B2:D6
      worksheet.addImage(chartImageId2, "B8:H20");
    }

    const workbookBuffer = await workbook.xlsx.writeBuffer();
    
    // send file to client
    saveAs(
      new Blob([workbookBuffer], { type: "application/octet-stream" }),
      `cecdata.xlsx`
    );
  };
  return (
    <div className="App">
      <header className="App-header">
        {renderLineChart}
        <button onClick={makeExcel}>Export Excel</button>
      </header>
    </div>
  );
}

// some fake data to make the line chart look decent
const data = [
  {
    name: "Page A",
    uv: 4000,
    pv: 2400,
    amt: 2400,
  },
  {
    name: "Page B",
    uv: 3000,
    pv: 1398,
    amt: 2210,
  },
  {
    name: "Page C",
    uv: 2000,
    pv: 9800,
    amt: 2290,
  },
  {
    name: "Page D",
    uv: 2780,
    pv: 3908,
    amt: 2000,
  },
  {
    name: "Page E",
    uv: 1890,
    pv: 4800,
    amt: 2181,
  },
  {
    name: "Page F",
    uv: 2390,
    pv: 3800,
    amt: 2500,
  },
  {
    name: "Page G",
    uv: 3490,
    pv: 4300,
    amt: 2100,
  },
];

const svgToPng = (svg: any, width: number, height: number) => {
  return new Promise<string>((resolve, reject) => {
    let canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");

    if (!ctx) throw Error("No canvas context found");

    // Set background to white
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    let xml = new XMLSerializer().serializeToString(svg);
    let dataUrl = "data:image/svg+xml;utf8," + encodeURIComponent(xml);
    let img = new Image(width, height);

    img.onload = () => {
      ctx.drawImage(img, 0, 0);
      let imageData = canvas.toDataURL("image/png", 1.0);
      resolve(imageData);
    };

    img.onerror = () => reject();

    img.src = dataUrl;
  });
};

export default App;
