var filePath = sessionStorage.getItem("filePath");
document.getElementById("fileName").innerText = filePath.split("\\").pop();