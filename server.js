const express = require("express");
const bodyParser = require("body-parser");
const fileUpload = require("express-fileupload");
const XLSX = require("xlsx");
const fs = require("fs");
const fetch = require("node-fetch");
require("dotenv").config();

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.use(
  fileUpload({
    createParentPath: true,
  })
);
app.post("/", async function (req, res) {
  try {
    if (!req.files) {
      res.send({
        status: false,
        message: "No file uploaded",
      });
    } else {
      try {
        const excelSheet = req.files.file;
        var fileUri = process.env.TEMP_SHEET_STORE_URI + excelSheet.name;
        await excelSheet.mv(fileUri);
        var workbook = XLSX.readFile(fileUri);
        var sheets = workbook.SheetNames;
        var worksheet = workbook.Sheets[sheets[0]];
        var ranges = worksheet["!ref"].split(":");
        var rowRange = ranges[1].substring(1);
        for (var i = 1; i < parseInt(rowRange); ++i) {
          var code_cell_ref = XLSX.utils.encode_cell({ c: 0, r: i });
          var price_cell_ref = XLSX.utils.encode_cell({ c: 1, r: i });
          if (!worksheet[code_cell_ref.toString()]) break;
          const url =
            process.env.PRODUCT_PRICE_ENDPOINT_URL +
            worksheet[code_cell_ref.toString()].v;

          const resp = await fetch(url);
          const body = await resp.json();
          worksheet[price_cell_ref.toString()] = {
            t: "s",
            v: body.data.price,
          };
        }
        await XLSX.writeFile(workbook, "./uploads/processed.xlsx");
        res.status(200).sendFile(__dirname + "/uploads/processed.xlsx");

        fs.unlinkSync(fileUri);
      } catch (error) {
        console.log(error);
      }
    }
  } catch (err) {
    res.status(500).send(err);
  }
});

app.get("/", function (req, res) {
  res.send("Hello, World!");
});

app.listen(process.env.SERVER_PORT, function () {
  console.log("Server is running on port 3000");
});
