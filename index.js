const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs').promises;
const { execFile } = require('child_process');

// ─── Экспорт в DOCX ──────────────────────────────────────────────────────────
// Основной путь — Pandoc (https://pandoc.org): проверенный конвертер
// Markdown + LaTeX-формул в DOCX с нативными формулами Word (Office Math).
// Ищем исполняемый файл pandoc: в PATH, затем в типичных каталогах Windows.
const PANDOC_CANDIDATES = [
    'pandoc',                                        // из PATH
    path.join(process.env.LOCALAPPDATA || '', 'Pandoc', 'pandoc.exe'),
    'C:\\Program Files\\Pandoc\\pandoc.exe',
    path.join(process.env.PROGRAMFILES || '', 'Pandoc', 'pandoc.exe')
];

function runPandoc(inputFile, outputFile) {
    return new Promise((resolve, reject) => {
        const tryNext = (i) => {
            if (i >= PANDOC_CANDIDATES.length) {
                reject(new Error('Pandoc не найден'));
                return;
            }
            execFile(PANDOC_CANDIDATES[i],
                ['--version'],
                { timeout: 5000, windowsHide: true },
                (err) => {
                    if (err) return tryNext(i + 1);          // этот кандидат недоступен
                    execFile(PANDOC_CANDIDATES[i],
                        [inputFile, '-f', 'markdown+tex_math_dollars+tex_math_single_backslash', '-t', 'docx', '-o', outputFile],
                        { timeout: 60000, windowsHide: true },
                        (err2, stderr) => {
                            if (err2) reject(new Error(`Pandoc: ${stderr || err2.message}`));
                            else resolve();
                        });
                });
        };
        tryNext(0);
    });
}

async function convertWithPandoc(markdownContent) {
    const tmpIn = path.join(os.tmpdir(), `mdv-${Date.now()}.md`);
    const tmpOut = path.join(os.tmpdir(), `mdv-${Date.now()}.docx`);
    try {
        await fs.writeFile(tmpIn, markdownContent, 'utf-8');
        await runPandoc(tmpIn, tmpOut);
        return await fs.readFile(tmpOut);
    } finally {
        fs.unlink(tmpIn).catch(() => {});
        fs.unlink(tmpOut).catch(() => {});
    }
}

// Резервный путь — встроенный конвертер (пакет "docx"). Загружается лениво:
// если пакет не установлен, приложение работает как раньше.
let docxExportModule = null;
function getDocxExporter() {
    if (!docxExportModule) {
        try {
            docxExportModule = require('./docx-export');
        } catch (err) {
            console.warn('Встроенный модуль экспорта DOCX не загружен:', err.message);
            return null;
        }
    }
    return docxExportModule;
}

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
    // Абсолютный путь — чтобы загрузка index.html не зависела от текущего каталога
    mainWindow.loadFile(path.join(__dirname, 'web', 'index.html'));
    
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
            { name: 'Word Document', extensions: ['docx'] },
            { name: 'All Files', extensions: ['*'] }
        ]
    });
    if (canceled || !filePath) return { canceled: true };

    const isDocx = /\.docx$/i.test(filePath);

    try {
        if (isDocx) {
            // 1) Основной путь: Pandoc — надёжный экспорт текста и формул (Office Math).
            let buffer = null;
            let pandocError = null;
            try {
                buffer = await convertWithPandoc(content);
            } catch (err) {
                pandocError = err.message;
                console.warn('Pandoc недоступен/ошибка, пробую встроенный конвертер:', pandocError);
            }

            // 2) Резервный путь: встроенный конвертер на пакете "docx".
            if (!buffer) {
                const exporter = getDocxExporter();
                if (!exporter) {
                    throw new Error(pandocError && pandocError.includes('Pandoc не найден')
                        ? 'Не найден Pandoc, а встроенный конвертер недоступен. Установите Pandoc (winget install --id JohnMacFarlane.Pandoc) либо выполните "npm install" в папке проекта.'
                        : `Не удалось создать DOCX: ${pandocError || 'встроенный модуль недоступен'}`);
                }
                buffer = await exporter.exportMarkdownToDocx(content);
            }
            await fs.writeFile(filePath, buffer);
        } else {
            await fs.writeFile(filePath, content, 'utf-8');
        }
        return { canceled: false, file_path: filePath, format: isDocx ? 'docx' : 'text' };
    } catch (err) {
        return { canceled: false, error: err.message };
    }
});