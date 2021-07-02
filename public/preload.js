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

  getWorkBook: () => {
    var filePath = sessionStorage.getItem("filePath");
    ipcRenderer.send("getWorkBook", filePath);
  },
});
