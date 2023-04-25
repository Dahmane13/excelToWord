var express = require("express");
var router = express.Router();
const XLSX = require("xlsx");

const fs = require("fs");

const ChartJSImage = require("chart.js-image");

const upload = require("../middlewares/multer-config");
const path = require("path");
const genDocx = require("../utils/genDocx");
/* GET home page. */
router.get("/", function (req, res, next) {
  const workbook = XLSX.readFile("./public/uploads/test.xlsx");
  const sheetName = workbook.SheetNames[0];

  // Get the worksheet object for the first sheet
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet);

  console.log(data[0]["sex"]);

  res.render("index", { title: "Express" });
});
router.post("/", upload.single("file"), async (req, res) => {
  const file = req.file;
  console.log(file.path);
  const workbook = XLSX.readFile(file.path);
  const sheetName = workbook.SheetNames[0];

  // Get the worksheet object for the first sheet
  const worksheet = workbook.Sheets[sheetName];
  let data = XLSX.utils.sheet_to_json(worksheet);
  data = data.sort(({ nmbr: a }, { nmbr: b }) => b - a);
  const total = data.find((item) => (item.sex = "total")).nmbr;
  const line_chart = ChartJSImage()
    .chart({
      type: "pie",
      data: {
        labels: [data[1]["sex"], data[2]["sex"]],
        datasets: [
          {
            label: "My First Dataset",
            data: [
              (data[1]["nmbr"] * 100) / total,
              (data[2]["nmbr"] * 100) / total,
            ],
            backgroundColor: ["rgb(255, 99, 132)", "rgb(54, 162, 235)"],
            hoverOffset: 4,
          },
        ],
      },
      options: {
        legend: {
          position: "bottom", // Change legend position to bottom
        },
      },
    }) // Line chart

    .backgroundColor("white")
    .width(500) // 500px
    .height(300); // 300px
  const pieImage = line_chart.toBuffer();
  const tempData = {
    first: data[1]["sex"],
    second: data[2]["sex"],
    firstPer: (data[1]["nmbr"] * 100) / total,
    secondPer: (data[2]["nmbr"] * 100) / total,
    image: pieImage,
  };
  const base64 = await genDocx(tempData);
  fs.writeFile("test2.docx", base64, { encoding: "base64" }, function (err) {
    console.log("File created");
  });
  res.setHeader("Content-Disposition", "attachment; filename=rapport.docx");
  res.send(Buffer.from(base64, "base64"));
});
module.exports = router;
