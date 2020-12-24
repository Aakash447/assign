const { default: axios } = require("axios");
const express = require("express");
const app = express();
const ExcelJS = require("exceljs");

// setting middlewares
app.set("view engine", "ejs");

app.get("/", async (req, res) => {
  try {
    let data1 = await axios.get("https://jsonplaceholder.typicode.com/users/");
    // console.log("data1:", data1.data);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("users");
    sheet.columns = [
      { header: "Name", key: "name", width: 30 },
      { header: "Username", key: "username", width: 30 },
      { header: "Email", key: "email", width: 30 },
      { header: "ZipCode", key: "zipcode", width: 30 },
    ];

    data1.data.map((row) => {
      sheet.addRow([row.name, row.username, row.email, row.address.zipcode]);
    });

    sheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true };
    });
    const data = await workbook.xlsx.writeFile("users.xlsx");
    res.send("excel file created..");
  } catch (err) {
    console.log("err:", err);
  }
});

app.get("/show", (req, res) => {
  let header = [];
  let data = {};
  const workbook = new ExcelJS.Workbook();
  workbook.xlsx.readFile("users.xlsx").then(function () {
    var worksheet = workbook.getWorksheet("users");
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      //   console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
      if (rowNumber == 1) {
        row.values.map((row, i) => {
          header.push(row);
        });
      } else {
        data = {
          ...data,
          [rowNumber]: row.values,
        };
      }
    });
    // console.log('header:',header)
    // console.log('data:',data)
  });

  setTimeout(() => {
    // console.log("header:", header);
    // console.log("data:", data);
    res.render("showTable", { header: header, data: data });
  }, 500);

  // res.render('showTable',{ header:header,data:data   })
});

const port = 3000;
app.listen(port, () => {
  console.log("Server is listening at port 3000");
});
