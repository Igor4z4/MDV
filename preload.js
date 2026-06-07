const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    openFileDialog: () => ipcRenderer.invoke('dialog:openFile'),
    readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
    onOpenFile: (callback) => ipcRenderer.on('open-file', (event, filePath) => callback(filePath))
});