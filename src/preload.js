const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('reimburse', {
  selectFiles: () => ipcRenderer.invoke('select-files'),
  selectOutputFolder: () => ipcRenderer.invoke('select-output-folder'),
  analyzeFiles: (filePaths) => ipcRenderer.invoke('analyze-files', filePaths),
  generatePackage: (payload) => ipcRenderer.invoke('generate-package', payload),
  onProgress: (callback) => ipcRenderer.on('generation-progress', (_event, payload) => callback(payload)),
});
