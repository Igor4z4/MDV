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
        autoHideMenuBar: true,   // скрываем стандартное меню Electron (Alt — показать)
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
        }
    });
    mainWindow.setMenuBarVisibility(false);
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

// IPC: открыть файл через диалог (кнопка "Открыть файл").
// Поддерживается выбор сразу нескольких файлов — каждый откроется в новой вкладке.
ipcMain.handle('dialog:openFile', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openFile', 'multiSelections'],
        filters: [
            { name: 'Markdown', extensions: ['md', 'markdown', 'txt'] },
            { name: 'All Files', extensions: ['*'] }
        ]
    });
    if (canceled || filePaths.length === 0) return { canceled: true };
    const files = [];
    for (const p of filePaths) {
        try {
            const content = await fs.readFile(p, 'utf-8');
            files.push({ file_path: p, content });
        } catch (err) {
            console.error('Ошибка чтения файла:', p, err);
        }
    }
    if (files.length === 0) return { canceled: true };
    // Обратная совместимость: file_path/content указывают на первый файл
    return { canceled: false, file_path: files[0].file_path, content: files[0].content, files };
});

// IPC: читать файл по пути (для открытия из аргументов командной строки)
ipcMain.handle('read-file', async (event, filePath) => {
    const content = await fs.readFile(filePath, 'utf-8');
    return content;
});

// IPC: сохранить содержимое в файл (Ctrl+S для файла, уже имеющего путь)
ipcMain.handle('save-file', async (event, { filePath, content }) => {
    try {
        await fs.writeFile(filePath, content, 'utf-8');
        return { ok: true };
    } catch (err) {
        return { ok: false, error: err.message };
    }
});

// IPC: сохранить как — показать диалог выбора пути
ipcMain.handle('save-file-as', async (event, { defaultPath, content }) => {
    const { canceled, filePath } = await dialog.showSaveDialog({
        defaultPath,
        filters: [
            { name: 'Markdown', extensions: ['md', 'markdown'] },
            { name: 'Text', extensions: ['txt'] },
            { name: 'All Files', extensions: ['*'] }
        ]
    });
    if (canceled || !filePath) return { canceled: true };
    try {
        await fs.writeFile(filePath, content, 'utf-8');
        return { canceled: false, file_path: filePath };
    } catch (err) {
        return { canceled: false, error: err.message };
    }
});