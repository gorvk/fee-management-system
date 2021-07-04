const electron = require("electron");
const url = require("url");
const path = require("path");
const { app, BrowserWindow, ipcMain, dialog } = electron;
const XLSX = require("exceljs");

let mainWindow;
global.filepath = undefined;

function initializeWorkbook(filePath) {
  var workbook = new XLSX.Workbook();
  workbook.xlsx.readFile(filePath).then(() => {
    if (workbook.getWorksheet("Fee") == undefined) {
      console.log("Worksheet Intialized !!");
      workbook.addWorksheet("Fee");
      workbook.xlsx.writeFile(filePath);
    } else {
      console.log("WORKSHEET Exists !!");
    }
  });
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
    .then((file) => {
      console.log(file.canceled);
      if (!file.canceled) {
        initializeWorkbook(file.filePaths[0]);
        event.reply("filePath", file.filePaths[0]);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on("delete", (event, data) => {
  // READ
  var found = {};
  var workbook = new XLSX.Workbook();
  workbook.xlsx.readFile(data.filePath).then(() => {
    var worksheet = workbook.getWorksheet("Fee");
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      if (
        row.values.includes(
          data.firstName,
          data.lastName,
          data.middleName,
          data.installment
        )
      ) {
        found = {
          firstName: row.values[1],
          middleName: row.values[2],
          lastName: row.values[3],
          mobileNumber : row.values[4],
          class: row.values[5],
          installment: row.values[6],
          paidAmount: row.values[7],
          billNumber: row.values[8],
        };
      }
    });
    event.reply("studentData", found);
  });
});

ipcMain.on("update", (event, data) => {
  // READ
  var found = {};
  var workbook = new XLSX.Workbook();
  console.log("Running Main Update");
  workbook.xlsx.readFile(data.filePath).then(() => {
    var worksheet = workbook.getWorksheet("Fee");
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      if (
        row.values.includes(
          data.firstName,
          data.lastName,
          data.middleName,
          data.installment
        )
      ) {
        found = {
          firstName: row.values[1],
          middleName: row.values[2],
          lastName: row.values[3],
          mobileNumber: row.values[4],
          class: row.values[5],
          installment: row.values[6],
          paidAmount: row.values[7],
          billNumber: row.values[8],
        };
      }
    });
    event.reply("studentData", found);
  });
});

ipcMain.on("insert", (event, data) => {
  var workbook = new XLSX.Workbook();
  workbook.xlsx.readFile(data.filePath).then(() => {
    var worksheet = workbook.getWorksheet("Fee");

    worksheet.columns = [
      { header: "First Name", key: "firstName", width: 10 },
      { header: "Middle Name", key: "middleName", width: 10 },
      { header: "Last Name", key: "lastName", width: 10 },
      { header: "Mobile Number", key: "mobileNumber", width: 10 },
      { header: "Class", key: "class", width: 10 },
      { header: "Installment", key: "installment", width: 10 },
      { header: "Paid Amount", key: "paidAmount", width: 10 },
      { header: "Bill Number", key: "billNumber", width: 10 },
    ];

    worksheet.addRow({
      firstName: data.firstName,
      middleName: data.middleName,
      lastName: data.lastName,
      mobileNumber: data.mobileNumber,
      class: data.class,
      installment: data.installment,
      paidAmount: data.paidAmount,
      billNumber: data.billNo,
    });

    workbook.xlsx.writeFile(data.filePath);
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
  });
  mainWindow.loadURL(
    url.format(path.join(__dirname, "public/index.html"), "file:", true)
  );
  mainWindow.maximize(true);
});
