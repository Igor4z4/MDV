const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs').promises;

function createWindow() {
    const win = new BrowserWindow({
        width: 1300,
        height: 850,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
        }
    });
    win.loadFile('web/index.html');
    // win.webContents.openDevTools(); // для отладки
}

// Обработчик диалога выбора файла
ipcMain.handle('dialog:openFile', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [
            { name: 'Markdown', extensions: ['md', 'markdown', 'txt'] },
            { name: 'All Files', extensions: ['*'] }
        ]
    });
    if (canceled) return { canceled: true };
    const content = await fs.readFile(filePaths[0], 'utf-8');
    return { canceled: false, file_path: filePaths[0], content };
});

app.whenReady().then(() => {
    createWindow();
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});