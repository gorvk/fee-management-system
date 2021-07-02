const electron = require("electron");
const url = require("url");
const path = require("path");
const { app, BrowserWindow, ipcMain, dialog } = electron;
const XLSX = require("exceljs");

let mainWindow;
global.filepath = undefined;

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
        global.filepath = file.filePaths[0];
        event.reply("filePath", global.filepath);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on("getWorkBook", async (event, data) => {
  var workbook = new XLSX.Workbook();
  workbook.xlsx.readFile(data).then(() => {
    var worksheet = workbook.getWorksheet("gourav");

    //READ
    // worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
    //   console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
    // });

    //WRITE
    // worksheet.addRow([1, 2, 3]);
    // workbook.xlsx.writeFile(data);
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
