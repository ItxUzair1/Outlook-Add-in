const { app, BrowserWindow } = require('electron');
const path = require('path');

// 1. Start the Node.js Express Backend
// This will start the API server on port 4001 and serve the React files
require('./src/server.js');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    title: 'Koyomail Admin Indexer',
    icon: path.join(__dirname, 'icon.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
    // Prevent the default Electron menu bar
    autoHideMenuBar: true,
  });

  // Wait briefly for the Express server to fully bind to the port
  setTimeout(() => {
    mainWindow.loadURL('http://localhost:4001');
  }, 1000);

  mainWindow.on('closed', function () {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});
