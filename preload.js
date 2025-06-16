// Electron APIs are available in the preload script
const { ipcRenderer, contextBridge } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('electronAPI', {
  // Dialogs
  openDialog: (options) => ipcRenderer.invoke('dialog:openFile', options),
  saveDialog: (options) => ipcRenderer.invoke('dialog:saveFile', options),
  // File System - Text Files
  readFile: (filePath, encoding = 'utf-8') => ipcRenderer.invoke('fs:readFile', filePath, encoding),
  writeFile: (filePath, data, encoding = 'utf-8') => ipcRenderer.invoke('fs:writeFile', filePath, data, encoding),
  // File System - Binary Files (like Excel, PDF)
  readBinaryFile: (filePath) => ipcRenderer.invoke('fs:readBinaryFile', filePath),
  writeBinaryFile: (filePath, data) => ipcRenderer.invoke('fs:writeBinaryFile', filePath, data),
  // Window operations
  openToolWindow: (toolPath) => ipcRenderer.send('open-tool', toolPath)
});

window.addEventListener('DOMContentLoaded', () => {
  // Example: You can add specific logic here if needed for all windows
  console.log('Preload script loaded.');
}); 