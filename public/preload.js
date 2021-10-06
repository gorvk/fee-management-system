// Preload Script. 

//Importing the ContextBridge for sharing the context data and IPCRenderer for sending and recieving data.
const { contextBridge, ipcRenderer } = require("electron"); 
var studentData;

// Function to perform Crud Operation on Excel File.
contextBridge.exposeInMainWorld("api", {
  getFilePath: () => {
    ipcRenderer.send("getFilePath");
    ipcRenderer.on("filePath", (_, data) => {
      sessionStorage.setItem("filePath", data);
      document.getElementById("btnGetFile").innerText = data.split("\\").pop();
      window.location.href = "./routes/insert.html";
    });
  },

  insert: () => {
    var filePath = sessionStorage.getItem("filePath");
    var form = document.getElementById("insertForm");
    var formData = new FormData(form);
    formData.append("filePath", filePath);
    var fdata = {};
    formData.forEach((value, key) => (fdata[key] = value));
    ipcRenderer.send("insert", fdata);
    
    ipcRenderer.on("fileError", (_, data) => {
      window.location.href = "../index.html";
    });
  },

  getData: () => {
    var filePath = sessionStorage.getItem("filePath");
    var formData = {
      firstName: document.getElementById("firstName").value.trim(),
      middleName: document.getElementById("middleName").value.trim(),
      lastName: document.getElementById("lastName").value.trim(),
      feeType: document.getElementById("feeType").value,
      filePath: filePath,
    };

    ipcRenderer.send("getData", formData);
    ipcRenderer.on("studentData", (_, data) => {
      if (data) {
        console.log("Data = " + data.date);
        studentData = data;
        studentData["filePath"] = filePath;
        document.getElementById("mobileNumber").value = data.mobileNumber;
        document.getElementById("class").value = data.class;
        document.getElementById("paidAmount").value = data.paidAmount;
        document.getElementById("billNo").value = data.billNo;
        document.getElementById("date").value = data.date;

        console.log("preload date = " + data.date);
        document.getElementById("submitButton").disabled = false;
      } else {
        console.log("SHOW DIALOG as -> Data Not Found");
      }
    });
  },

  update: () => {
    var filePath = sessionStorage.getItem("filePath");
    var form = document.getElementById("updateForm");
    var formData = new FormData(form);
    formData.append("filePath", filePath);
    formData.append("rowNumber", studentData.rowNumber);
    var fdata = {};
    formData.forEach((value, key) => (fdata[key] = value));
    ipcRenderer.send("update", fdata);
  },

  delete: () => {
    ipcRenderer.send("delete", studentData);
  },
});
