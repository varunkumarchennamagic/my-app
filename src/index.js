const { app, BrowserWindow, dialog } = require("electron");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");
const { ipcMain } = require('electron');

// Listen for the file-upload event from the renderer process
ipcMain.on('file-upload', () => {
  handleFileUpload();
});

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      nodeIntegration: true,
    },
  });

  // and load the index.html of the app.
  mainWindow.loadFile(path.join(__dirname, "index.html"));

  // Open the DevTools.
  mainWindow.webContents.openDevTools();

  mainWindow.on("closed", function () {
    mainWindow = null;
  });
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on("ready", createWindow);

app.on("window-all-closed", function () {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", function () {
  if (mainWindow === null) {
    createWindow();
  }
});

// Function to handle file upload and conversion
function handleFileUpload() {
  dialog
    .showOpenDialog(mainWindow, {
      properties: ["openFile"],
      filters: [{ name: "XLSX Files", extensions: ["xlsx"] }],
    })
    .then((result) => {
      if (!result.canceled) {
        const filePath = result.filePaths[0];
        const workbook = xlsx.readFile(filePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = xlsx.utils.sheet_to_json(worksheet);
        const jsonFilePath = filePath.replace(".xlsx", ".json");

        fs.writeFile(jsonFilePath, JSON.stringify(jsonData, null, 2), (err) => {
          if (err) {
            dialog.showErrorBox(
              "Error",
              "An error occurred while converting the file."
            );
          } else {
            dialog.showMessageBox(mainWindow, {
              message: "File converted successfully!",
              buttons: ["OK"],
            });
          }
        });
      }
    });
}

// Expose handleFileUpload function to the renderer process
// You can call this function from your HTML/JavaScript code
global.handleFileUpload = handleFileUpload;
