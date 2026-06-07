const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs').promises;

let mainWindow = null;
let pendingFilePath = null;

// Функция создания окна
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1300,
        height: 850,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
        }
    });
    mainWindow.loadFile('web/index.html');
    
    // Если есть файл, который нужно открыть после загрузки окна
    if (pendingFilePath) {
        mainWindow.webContents.once('did-finish-load', () => {
            mainWindow.webContents.send('open-file', pendingFilePath);
            pendingFilePath = null;
        });
    }
}

// Обработка одиночного экземпляра приложения (чтобы при повторном открытии файла не создавалось новое окно)
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
    app.quit();
} else {
    app.on('second-instance', (event, commandLine, workingDirectory) => {
        // Ищем аргумент, похожий на путь к .md файлу
        const filePath = commandLine.find(arg => 
            arg.endsWith('.md') || arg.endsWith('.markdown') || arg.endsWith('.txt')
        );
        if (filePath && mainWindow) {
            if (mainWindow.isMinimized()) mainWindow.restore();
            mainWindow.focus();
            mainWindow.webContents.send('open-file', filePath);
        }
    });

    app.whenReady().then(() => {
        // Проверяем аргументы при первом запуске
        const args = process.argv;
        for (const arg of args) {
            if (arg !== process.execPath && arg !== __filename && 
                (arg.endsWith('.md') || arg.endsWith('.markdown') || arg.endsWith('.txt'))) {
                pendingFilePath = arg;
                break;
            }
        }
        createWindow();
        app.on('activate', () => {
            if (BrowserWindow.getAllWindows().length === 0) createWindow();
        });
    });
}

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

// IPC: открыть файл через диалог (кнопка "Открыть файл")
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

// IPC: читать файл по пути (для открытия из аргументов командной строки)
ipcMain.handle('read-file', async (event, filePath) => {
    const content = await fs.readFile(filePath, 'utf-8');
    return content;
});