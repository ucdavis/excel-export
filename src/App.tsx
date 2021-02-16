import React, { useRef } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { LineChart, Line } from "recharts";

import "./App.css";

function App() {
  const chartRef = useRef<any>(null);

  const renderLineChart = (
    <LineChart ref={chartRef} width={400} height={400} data={chartData}>
      <Line type="monotone" dataKey="uv" stroke="#8884d8" />
    </LineChart>
  );

  const makeExcel = async () => {
    // https://github.com/exceljs/exceljs#interface
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("ExcelJS sheet");

    // const row = worksheet.getRow(5);
    // row.getCell(1).value = 5; // A5's value set to 5
    // row.getCell("C").value = new Date(); // C5's value set to now

    const yearlyHeaders = [];

    for (let i = 0; i < 30; i++) {
      yearlyHeaders.push({ name: "Y" + (new Date().getFullYear() + i) });
    }

    worksheet.addTable({
      name: "TechPerf",
      ref: "B2",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Technical Performance" }, { name: " " }],
      rows: [
        ["Project Prescription", "Clearcut"],
        ["Facility Type", "Bio Fuel"],
        ["Capital Cost ($)", 123.123],
      ],
    });

    worksheet.addTable({
      name: "supply",
      ref: "B12",
      headerRow: true,
      totalsRow: false,
      columns: [
        { name: "Resource Supply (ton)" },
        { name: "Total" },
        ...yearlyHeaders,
      ],
      rows: [
        [
          "Feedstock ",
          data.Resources.Feedstock.Total,
          ...data.Resources.Feedstock.Yearly,
        ],
        [
          "Coproduct",
          data.Resources.Coproduct.Total,
          ...data.Resources.Coproduct.Yearly,
        ],
      ],
    });

    worksheet.addTable({
      name: "analysis",
      ref: "B16",
      headerRow: true,
      totalsRow: false,
      columns: [
        { name: "Environmental Analysis" },
        { name: "Unit" },
        { name: "Total" },
        ...yearlyHeaders,
      ],
      rows: [
        [
          "Diesel",
          data.Environemntal.Diesel.Unit,
          data.Environemntal.Diesel.Total,
          ...data.Environemntal.Diesel.Yearly,
        ],
        [
          "Gasoline",
          data.Environemntal.Gasoline.Unit,
          data.Environemntal.Gasoline.Total,
          ...data.Environemntal.Gasoline.Yearly,
        ],
        [
          "Jet Fuel",
          data.Environemntal.JetFuel.Unit,
          data.Environemntal.JetFuel.Total,
          ...data.Environemntal.JetFuel.Yearly,
        ],
        [
          "Transport Distance",
          data.Environemntal.TransportDistance.Unit,
          data.Environemntal.TransportDistance.Total,
          ...data.Environemntal.TransportDistance.Yearly,
        ],
      ],
    });

    worksheet.addTable({
      name: "lci",
      ref: "B22",
      headerRow: true,
      totalsRow: false,
      columns: [
        { name: "LCI Results" },
        { name: "Unit" },
        { name: "Total" },
        ...yearlyHeaders,
      ],
      rows: [
        [
          "CO2",
          data.LCI.CO2.Unit,
          data.LCI.CO2.Total,
          ...data.LCI.CO2.Yearly,
        ],
        [
          "CH4",
          data.LCI.CH4.Unit,
          data.LCI.CH4.Total,
          ...data.LCI.CH4.Yearly,
        ],
        [
          "N2O",
          data.LCI.N2O.Unit,
          data.LCI.N2O.Total,
          ...data.LCI.N2O.Yearly,
        ],
        [
          "CO2e",
          data.LCI.CO2e.Unit,
          data.LCI.CO2e.Total,
          ...data.LCI.CO2e.Yearly,
        ],
        [
          "CO",
          data.LCI.CO.Unit,
          data.LCI.CO.Total,
          ...data.LCI.CO.Yearly,
        ],
        [
          "NOx",
          data.LCI.NOx.Unit,
          data.LCI.NOx.Total,
          ...data.LCI.NOx.Yearly,
        ],
        [
          "NH3",
          data.LCI.NH3.Unit,
          data.LCI.NH3.Total,
          ...data.LCI.NH3.Yearly,
        ],
        [
          "PM10",
          data.LCI.CO2.Unit,
          data.LCI.CO2.Total,
          ...data.LCI.CO2.Yearly,
        ],
        [
          "PM2.5",
          data.LCI.PM25.Unit,
          data.LCI.PM25.Total,
          ...data.LCI.PM25.Yearly,
        ],
        [
          "SO2",
          data.LCI.SO2.Unit,
          data.LCI.SO2.Total,
          ...data.LCI.SO2.Yearly,
        ],
        [
          "SOx",
          data.LCI.SOx.Unit,
          data.LCI.SOx.Total,
          ...data.LCI.SOx.Yearly,
        ],
        [
          "VOCs",
          data.LCI.VOCs.Unit,
          data.LCI.VOCs.Total,
          ...data.LCI.VOCs.Yearly,
        ],
        [
          "Carbon Intensity",
          data.LCI.CarbonIntensity.Unit,
          data.LCI.CarbonIntensity.Total,
          ...data.LCI.CarbonIntensity.Yearly,
        ],
      ],
    });

    worksheet.addTable({
      name: "lcia",
      ref: "B37",
      headerRow: true,
      totalsRow: false,
      columns: [
        { name: "LCIA Results" },
        { name: "Unit" },
        { name: "Total" },
        ...yearlyHeaders,
      ],
      rows: [
        [
          "Global Warming Air",
          data.LCIA.GlobalWarmingAir.Unit,
          data.LCIA.GlobalWarmingAir.Total,
          ...data.LCIA.GlobalWarmingAir.Yearly,
        ],
        [
          "Acidification Air",
          data.LCIA.AcidificationAir.Unit,
          data.LCIA.AcidificationAir.Total,
          ...data.LCIA.AcidificationAir.Yearly,
        ],
        [
          "HH Particulate Air",
          data.LCIA.HHParticulateAir.Unit,
          data.LCIA.HHParticulateAir.Total,
          ...data.LCIA.HHParticulateAir.Yearly,
        ],
        [
          "Euthrophication Air",
          data.LCIA.EuthrophicationAir.Unit,
          data.LCIA.EuthrophicationAir.Total,
          ...data.LCIA.EuthrophicationAir.Yearly,
        ],
        [
          "Euthrophication Water",
          data.LCIA.EuthrophicationWater.Unit,
          data.LCIA.EuthrophicationWater.Total,
          ...data.LCIA.EuthrophicationWater.Yearly,
        ],
        [
          "Smog Air",
          data.LCIA.SmogAir.Unit,
          data.LCIA.SmogAir.Total,
          ...data.LCIA.SmogAir.Yearly,
        ],
      ]
    })

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
      worksheet.addImage(chartImageId2, "B75:J100");
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

const fakeYearlyData = () =>
  Array.from({ length: 30 }, () => Math.floor(Math.random() * 40));

const data = {
  Resources: {
    Feedstock: {
      Total: 123,
      Yearly: fakeYearlyData(),
    },
    Coproduct: {
      Total: 321,
      Yearly: fakeYearlyData(),
    },
  },
  Environemntal: {
    Diesel: {
      Unit: "gal",
      Total: 123,
      Yearly: fakeYearlyData(),
    },
    Gasoline: {
      Unit: "gal",
      Total: 456,
      Yearly: fakeYearlyData(),
    },
    JetFuel: {
      Unit: "gal",
      Total: 789,
      Yearly: fakeYearlyData(),
    },
    TransportDistance: {
      Unit: "km",
      Total: 101112,
      Yearly: fakeYearlyData(),
    },
  },
  LCI: {
    CO2: {
      Unit: "kg",
      Total: 123,
      Yearly: fakeYearlyData(),
    },
    CH4: {
      Unit: "g",
      Total: 235,
      Yearly: fakeYearlyData(),
    },
    N2O: {
      Unit: "g",
      Total: 364,
      Yearly: fakeYearlyData(),
    },
    CO2e: {
      Unit: "kg",
      Total: 785,
      Yearly: fakeYearlyData(),
    },
    CO: {
      Unit: "kg",
      Total: 918,
      Yearly: fakeYearlyData(),
    },
    NOx: {
      Unit: "g",
      Total: 534,
      Yearly: fakeYearlyData(),
    },
    NH3: {
      Unit: "g",
      Total: 123,
      Yearly: fakeYearlyData(),
    },
    PM10: {
      Unit: "g",
      Total: 324,
      Yearly: fakeYearlyData(),
    },
    PM25: {
      Unit: "g",
      Total: 123,
      Yearly: fakeYearlyData(),
    },
    SO2: {
      Unit: "g",
      Total: 903,
      Yearly: fakeYearlyData(),
    },
    SOx: {
      Unit: "g",
      Total: 234,
      Yearly: fakeYearlyData(),
    },
    VOCs: {
      Unit: "g",
      Total: 535,
      Yearly: fakeYearlyData(),
    },
    CarbonIntensity: {
      Unit: "g",
      Total: 234,
      Yearly: fakeYearlyData(),
    },
  },
  LCIA: {
    GlobalWarmingAir: {
      Unit: "kg CO2 eq",
      Total: 234,
      Yearly: fakeYearlyData(),
    },
    AcidificationAir: {
      Unit: "kg SO2 eq",
      Total: 234,
      Yearly: fakeYearlyData(),
    },
    HHParticulateAir: {
      Unit: "PM2.5 eq",
      Total: 234,
      Yearly: fakeYearlyData(),
    },
    EuthrophicationAir: {
      Unit: "kg N eq",
      Total: 234,
      Yearly: fakeYearlyData(),
    },
    EuthrophicationWater: {
      Unit: "kg N eq",
      Total: 234,
      Yearly: fakeYearlyData(),
    },
    SmogAir: {
      Unit: "kg O3 eq",
      Total: 234,
      Yearly: fakeYearlyData(),
    },
  }
};

// some fake data to make the line chart look decent
const chartData = [
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
