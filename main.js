const electron = require("electron");
const url = require("url");
const path = require("path");
const { app, BrowserWindow, ipcMain, dialog } = electron;
const XLSX = require("exceljs");

let mainWindow;

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

ipcMain.on("getData", (event, data) => {
  // READ
  var found = false;
  var workbook = new XLSX.Workbook();
  workbook.xlsx.readFile(data.filePath).then(() => {
    var worksheet = workbook.getWorksheet("Fee");
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      if (
        row.values.includes(data.firstName) &&
        row.values.includes(data.lastName) &&
        row.values.includes(data.middleName) &&
        row.values.includes(data.installment)
      ) {
        found = {
          mobileNumber: row.values[4],
          class: row.values[5],
          paidAmount: row.values[7],
          billNumber: row.values[8],
          rowNumber: rowNumber,
        };
      }
    });
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
    installment: 6,
    paidAmount: 7,
    billNumber: 8,
  };
  workbook.xlsx.readFile(data.filePath).then(() => {
    var worksheet = workbook.getWorksheet("Fee");
    var row = worksheet.getRow(data.rowNumber);
    row.getCell(columns.firstName).value = data.firstName;
    row.getCell(columns.middleName).value = data.middleName;
    row.getCell(columns.lastName).value = data.lastName;
    row.getCell(columns.mobileNumber).value = data.mobileNumber;
    row.getCell(columns.class).value = data.class;
    row.getCell(columns.installment).value = data.installment;
    row.getCell(columns.paidAmount).value = data.paidAmount;
    row.getCell(columns.billNumber).value = data.billNumber;

    row.commit();
    workbook.xlsx.writeFile(data.filePath);
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
      installment: 6,
      paidAmount: 7,
      billNumber: 8,
    };
    var worksheet = workbook.getWorksheet("Fee");
    var row = worksheet.getRow(data.rowNumber);
    row.getCell(columns.firstName).value = "";
    row.getCell(columns.middleName).value = "";
    row.getCell(columns.lastName).value = "";
    row.getCell(columns.mobileNumber).value = "";
    row.getCell(columns.class).value = "";
    row.getCell(columns.installment).value = "";
    row.getCell(columns.paidAmount).value = "";
    row.getCell(columns.billNumber).value = "";
    row.commit();

    if (data.rowNumber == worksheet.rowCount) {
    } else {
      worksheet.spliceRows(data.rowNumber, 1);
    }
    workbook.xlsx.writeFile(data.filePath);
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
