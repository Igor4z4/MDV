const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    openFileDialog: () => ipcRenderer.invoke('dialog:openFile'),
    readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
    onOpenFile: (callback) => ipcRenderer.on('open-file', (event, filePath) => callback(filePath)),
    saveFile: (filePath, content) => ipcRenderer.invoke('save-file', { filePath, content }),
    saveFileAs: (defaultPath, content) => ipcRenderer.invoke('save-file-as', { defaultPath, content })
});