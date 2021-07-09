const electron = require("electron");
const url = require("url");
const path = require("path");
const { app, BrowserWindow, ipcMain, dialog } = electron;
const XLSX = require("exceljs");
const fs = require("fs");

let mainWindow;
let fp;

function initializeWorkbook(filePath) {
  let workbook = new XLSX.Workbook();

  workbook.xlsx
    .readFile(filePath)
    .then((wb) => {
      var worksheet = wb.getWorksheet("Fee");
      if (worksheet == undefined) {
        wb.addWorksheet("Fee");
        wb.xlsx.writeFile(filePath); // TypeError: Cannot read property 'date1904' of undefined
      } else {
        console.log("WORKSHEET Exists !!");
      }
    })
    .catch((err) => {
      console.log("Before Add Sheet 2");
      worksheet = workbook.addWorksheet("Fee");
      workbook.xlsx.writeFile(filePath);
    });
}

function showDialog(title, message) {
  var options = {
    message: message,
    title: title,
  };
  dialog.showMessageBox(null, options);
}

ipcMain.on("getFilePath", (event, data) => {
  dialog
    .showOpenDialog({
      filters: [
        {
          name: "Excel Files",
          extensions: ["xls", "xlsx"],
        },
      ],
      properties: ["openFile"],
    })
    .then(async (file) => {
      fp = file.filePaths[0];
      console.log(file.canceled);
      if (!file.canceled) {
        initializeWorkbook(fp);
        event.reply("filePath", fp);
      }
    })
    .catch((err) => {
      console.log("yo2");
      console.log(err);
    });
});

ipcMain.on("getData", (event, data) => {
  // READ
  var found = false;
  var workbook = new XLSX.Workbook();
  workbook.xlsx.readFile(data.filePath).then(() => {
    var worksheet = workbook.getWorksheet("Fee");
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      // console.log(row.values);
      if (
        row.values.includes(data.firstName) &&
        row.values.includes(data.lastName) &&
        row.values.includes(data.middleName) &&
        row.values.includes(data.feeType)
      ) {
        console.log("date main =" + row.values[9]);
        found = {
          mobileNumber: row.values[4],
          class: row.values[5],
          paidAmount: row.values[7],
          billNo: row.values[8],
          date: row.values[9],
          rowNumber: rowNumber,
        };
      }
    });

    if (found == false) {
      showDialog("Failed", "Data not Found");
    }
    event.reply("studentData", found);
  });
});

ipcMain.on("update", (event, data) => {
  var workbook = new XLSX.Workbook();
  var columns = {
    firstName: 1,
    middleName: 2,
    lastName: 3,
    mobileNumber: 4,
    class: 5,
    feeType: 6,
    paidAmount: 7,
    billNo: 8,
    date: 9,
  };
  workbook.xlsx.readFile(data.filePath).then(() => {
    var worksheet = workbook.getWorksheet("Fee");
    var row = worksheet.getRow(data.rowNumber);
    row.getCell(columns.firstName).value = data.firstName.trim();
    row.getCell(columns.middleName).value = data.middleName.trim();
    row.getCell(columns.lastName).value = data.lastName.trim();
    row.getCell(columns.mobileNumber).value = data.mobileNumber;
    row.getCell(columns.class).value = data.class;
    row.getCell(columns.feeType).value = data.feeType;
    row.getCell(columns.paidAmount).value = data.paidAmount;
    row.getCell(columns.billNo).value = data.billNo;
    row.getCell(columns.date).value = data.date;

    row.commit();
    workbook.xlsx.writeFile(data.filePath);
    showDialog("Success", "Data Updated");
  });
});
ipcMain.on("delete", (event, data) => {
  var workbook = new XLSX.Workbook();

  workbook.xlsx.readFile(data.filePath).then(() => {
    var columns = {
      firstName: 1,
      middleName: 2,
      lastName: 3,
      mobileNumber: 4,
      class: 5,
      feeType: 6,
      paidAmount: 7,
      billNo: 8,
      date: 9,
    };
    var worksheet = workbook.getWorksheet("Fee");
    var row = worksheet.getRow(data.rowNumber);
    row.getCell(columns.firstName).value = "";
    row.getCell(columns.middleName).value = "";
    row.getCell(columns.lastName).value = "";
    row.getCell(columns.mobileNumber).value = "";
    row.getCell(columns.class).value = "";
    row.getCell(columns.feeType).value = "";
    row.getCell(columns.paidAmount).value = "";
    row.getCell(columns.billNo).value = "";
    row.getCell(columns.date).value = "";
    row.commit();

    if (data.rowNumber == worksheet.rowCount) {
    } else {
      worksheet.spliceRows(data.rowNumber, 1);
    }
    workbook.xlsx.writeFile(data.filePath);
    showDialog("Success", "Data Deleted");
  });
});

ipcMain.on("insert", (event, data) => {
  var workbook = new XLSX.Workbook();

  workbook.xlsx
    .readFile(data.filePath)
    .then(() => {
      var worksheet = workbook.getWorksheet("Fee");

      worksheet.columns = [
        { header: "First Name", key: "firstName", width: 10 },
        { header: "Middle Name", key: "middleName", width: 10 },
        { header: "Last Name", key: "lastName", width: 10 },
        { header: "Mobile Number", key: "mobileNumber", width: 10 },
        { header: "Class", key: "class", width: 10 },
        { header: "Fee Type", key: "feeType", width: 10 },
        { header: "Paid Amount", key: "paidAmount", width: 10 },
        { header: "Bill Number", key: "billNo", width: 10 },
        { header: "Date", key: "date", width: 10 },
      ];

      worksheet.addRow({
        firstName: data.firstName.trim(),
        middleName: data.middleName.trim(),
        lastName: data.lastName.trim(),
        mobileNumber: data.mobileNumber,
        class: data.class,
        feeType: data.feeType,
        paidAmount: data.paidAmount,
        billNo: data.billNo,
        date: data.date,
      });
      workbook.xlsx.writeFile(data.filePath);
      showDialog("Success", "Data Inserted");
    })
    .catch((err) => {
      let flag = 1;
      dialog
        .showMessageBox(null, {
          message: "An Error Occured please select the file again.",
          title: "Error",
          buttons: ["OK"],
          defaultId: 1,
        })
        .then((result) => {
          if (result.response === 0) {
            flag = result.response;
            console.log("Responsed");
          }
        });
      event.reply("fileError");
    });
});

app.on("ready", () => {
  mainWindow = new BrowserWindow({
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      enableRemoteModule: true,
      preload: path.join(__dirname, "./public/preload.js"),
    },
    icon: "./public/styles/ic_launcher.jpg",
  });
  mainWindow.loadURL(
    url.format(path.join(__dirname, "public/index.html"), "file:", true)
  );
  mainWindow.maximize(true);
  mainWindow.removeMenu()
  
});
