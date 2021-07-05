const { contextBridge } = require("electron");
const { ipcRenderer } = require("electron");

var studentData;

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
  },

  getData: () => {
    var filePath = sessionStorage.getItem("filePath");
    var formData = {
      firstName: document.getElementById("firstName").value,
      middleName: document.getElementById("middleName").value,
      lastName: document.getElementById("lastName").value,
      installment: document.getElementById("installment").value,
      filePath: filePath,
    };

    ipcRenderer.send("getData", formData);
    ipcRenderer.on("studentData", (_, data) => {
      if (data) {
        studentData = data;
        studentData["filePath"] = filePath
        document.getElementById("mobileNumber").value = data.mobileNumber;
        document.getElementById("class").value = data.class;
        document.getElementById("paidAmount").value = data.paidAmount;
        document.getElementById("billNo").value = data.billNumber;
      } else {
        // SHOW DIALOG as -> Data Not Found
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
