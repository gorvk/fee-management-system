const { contextBridge } = require("electron");
const { ipcRenderer } = require("electron");

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

  update: () => {
    var fdata = {
      firstName: document.getElementById("firstName").value,
      middleName: document.getElementById("middleName").value,
      lastName: document.getElementById("lastName").value,
      installment: document.getElementById("installment").value,
      filePath: sessionStorage.getItem("filePath"),
    };

    ipcRenderer.send("update", fdata);
    ipcRenderer.on("studentData", (_, data) => {
      document.getElementById("mobileNumber").value = data.mobileNumber;
      document.getElementById("class").value = data.class;
      document.getElementById("paidAmount").value = data.paidAmount;
      document.getElementById("billNo").value = data.billNumber;
    });
  },

  delete: () => {
    var fdata = {
      firstName: document.getElementById("firstName").value,
      middleName: document.getElementById("middleName").value,
      lastName: document.getElementById("lastName").value,
      installment: document.getElementById("installment").value,
      filePath: sessionStorage.getItem("filePath"),
    };

    ipcRenderer.send("delete", fdata);
    ipcRenderer.on("studentData", (_, data) => {
      document.getElementById("mobileNumber").value = data.mobileNumber;
      document.getElementById("class").value = data.class;
      document.getElementById("paidAmount").value = data.paidAmount;
      document.getElementById("billNo").value = data.billNumber;
    });
  },
});
