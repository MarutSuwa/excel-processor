const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electron", {
  selectFile: () => ipcRenderer.invoke("select-file"),
  saveFile: (defaultPath) => ipcRenderer.invoke("save-file", defaultPath),
  processFile: (filePath, savePath, openTimeColumn, closeTimeColumn) =>
    ipcRenderer.invoke(
      "process-file",
      filePath,
      savePath,
      openTimeColumn,
      closeTimeColumn
    ),
});
