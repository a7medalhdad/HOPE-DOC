const { app, BrowserWindow, ipcMain, dialog, globalShortcut, Menu } = require('electron');
const { autoUpdater } = require('electron-updater');
const path = require('path');
const fs = require('fs');
const os = require('os');
// قراءة رمز GitHub من ملف التكوين
let githubToken = '';
try {
    const configPath = path.join(__dirname, 'config.json');
    if (fs.existsSync(configPath)) {
        const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
        githubToken = config.githubToken;
        if (githubToken && githubToken !== 'your_github_token_here') {
            process.env.GH_TOKEN = githubToken;
        }
    } else {
        // إنشاء ملف التكوين إذا لم يكن موجوداً
        fs.writeFileSync(configPath, JSON.stringify({ githubToken: 'your_github_token_here' }, null, 4));
        console.log('تم إنشاء ملف config.json. الرجاء تعديله وإضافة رمز GitHub الخاص بك.');
    }
} catch (error) {
    console.error('خطأ في قراءة ملف التكوين:', error);
}

let windowStateKeeper;
try {
    windowStateKeeper = require('electron-window-state');
} catch (error) {
    console.error('Failed to load electron-window-state:', error);
    // Fallback window state management
    windowStateKeeper = () => ({
        x: undefined,
        y: undefined,
        width: 1024,
        height: 768,
        manage: () => {},
        unmanage: () => {}
    });
}
const Store = require('electron-store');
const ExcelProcessor = require('./excel-processor');

// إنشاء ffmpeg.dll وهمي لحل مشاكل البيئات المحمية
try {
    require('./create-dummy-ffmpeg');
} catch (error) {
    console.log('Could not create dummy ffmpeg, continuing...');
}

// معالجة الأخطاء العامة للتطبيق
process.on('uncaughtException', (error) => {
    console.error('Uncaught Exception:', error);
    // لا نغلق التطبيق، فقط نسجل الخطأ
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection at:', promise, 'reason:', reason);
    // لا نغلق التطبيق، فقط نسجل الخطأ
});

// تعطيل التحقق من الشهادات للبيئات المحمية
app.commandLine.appendSwitch('ignore-certificate-errors');
app.commandLine.appendSwitch('ignore-ssl-errors');
app.commandLine.appendSwitch('ignore-certificate-errors-spki-list');
app.commandLine.appendSwitch('disable-features', 'OutOfBlinkCors');

// تعطيل ffmpeg وجميع الميزات المتعلقة بالوسائط
app.commandLine.appendSwitch('disable-features', 'MediaFoundationVideoCapture');
app.commandLine.appendSwitch('disable-features', 'HardwareMediaKeyHandling');
app.commandLine.appendSwitch('disable-background-media-suspend');
app.commandLine.appendSwitch('disable-renderer-backgrounding');
app.commandLine.appendSwitch('disable-backgrounding-occluded-windows');
app.commandLine.appendSwitch('disable-features', 'TranslateUI');
app.commandLine.appendSwitch('disable-features', 'VizDisplayCompositor');
app.commandLine.appendSwitch('disable-dev-shm-usage');
app.commandLine.appendSwitch('disable-accelerated-2d-canvas');
app.commandLine.appendSwitch('disable-accelerated-jpeg-decoding');
app.commandLine.appendSwitch('disable-accelerated-mjpeg-decode');
app.commandLine.appendSwitch('disable-accelerated-video-decode');
app.commandLine.appendSwitch('disable-accelerated-video-encode');

// تكوين النظام الهادئ
const QUIET_MODE = true; // تفعيل الوضع الهادئ
const SHOW_ERRORS_ONLY = true; // إظهار الأخطاء فقط

// دالة للطباعة الهادئة
function quietLog(...args) {
    if (!QUIET_MODE) {
        console.log(...args);
    }
}

// دالة للأخطاء المهمة فقط
function quietError(...args) {
    if (SHOW_ERRORS_ONLY) {
        console.error(...args);
    }
}

// دالة للتحذيرات المهمة فقط
function quietWarn(...args) {
    if (SHOW_ERRORS_ONLY) {
        console.warn(...args);
    }
}

// تحسين الأداء بإضافة معاملات command line (بدون GPU acceleration)
app.commandLine.appendSwitch('disable-renderer-backgrounding');
app.commandLine.appendSwitch('disable-gpu-sandbox');
app.commandLine.appendSwitch('disable-gpu');
app.commandLine.appendSwitch('disable-software-rasterizer');
app.commandLine.appendSwitch('disable-features', 'VizDisplayCompositor');
app.commandLine.appendSwitch('no-sandbox');
app.commandLine.appendSwitch('disable-web-security');
// حل مشكلة الكاش
app.commandLine.appendSwitch('disk-cache-size', '0');
app.commandLine.appendSwitch('media-cache-size', '0');

// تعيين مسار كاش مخصص في temp directory
const customCachePath = path.join(os.tmpdir(), 'hope-doc-cache');
try {
  // إنشاء المجلد إذا لم يكن موجوداً
  if (!fs.existsSync(customCachePath)) {
    fs.mkdirSync(customCachePath, { recursive: true });
  }
  app.setPath('userData', customCachePath);
} catch (error) {
      quietWarn('Failed to set custom cache path:', error);
}

// Importar módulos adicionales para la generación de documentos Word
let docx;
try {
  docx = require('docx');
} catch (error) {
      quietWarn('docx library not available:', error.message);
}

// Create settings storage
const store = new Store({ name: 'window-state' });

// Clave para almacenar la última ruta utilizada
const LAST_USED_PATH_KEY = 'lastUsedPath';

let mainWindow;
let toolWindows = new Map();
let minimizedState = false;
let cacheCleanupInterval;
let shortcutsConfigWindow;

// تكوين التحديث التلقائي
function configureAutoUpdater() {
    // تكوين التحديث التلقائي مع المصادقة
    if (process.env.GH_TOKEN) {
        autoUpdater.setFeedURL({
            provider: 'github',
            owner: 'ahmedalhddad',
            repo: 'HOPE-DOC',
            private: true,
            token: process.env.GH_TOKEN
        });
    }

    // تعطيل التحميل التلقائي للتحديثات
    autoUpdater.autoDownload = false;
    autoUpdater.allowDowngrade = false;

    // إضافة معالج حدث النقر على الشعار للتحقق من التحديثات
    ipcMain.on('check-for-updates', () => {
        if (!process.env.GH_TOKEN) {
            dialog.showMessageBox({
                type: 'error',
                title: 'خطأ في التحديث',
                message: 'تعذر التحقق من التحديثات: لم يتم تكوين رمز GitHub',
                buttons: ['حسناً']
            });
            return;
        }

        autoUpdater.checkForUpdates().catch(err => {
            dialog.showMessageBox({
                type: 'error',
                title: 'خطأ في التحديث',
                message: 'حدث خطأ أثناء التحقق من التحديثات',
                detail: err.message,
                buttons: ['حسناً']
            });
        });
    });

    // معالجة الأحداث المختلفة للتحديث
    autoUpdater.on('checking-for-update', () => {
        // إظهار نافذة منبثقة تخبر المستخدم أنه يتم التحقق من التحديثات
        dialog.showMessageBox({
            type: 'info',
            title: 'التحقق من التحديثات',
            message: 'جاري التحقق من وجود تحديثات جديدة...',
            buttons: ['حسناً']
        });
    });

    autoUpdater.on('update-available', (info) => {
        // إظهار نافذة منبثقة تخبر المستخدم بوجود تحديث جديد
        dialog.showMessageBox({
            type: 'info',
            title: 'تحديث متوفر',
            message: `يتوفر إصدار جديد (${info.version}).\nالإصدار الحالي: ${app.getVersion()}\n\nهل تريد تحميل التحديث الآن؟`,
            detail: 'سيتم تحميل التحديث في الخلفية وسيتم إعلامك عند اكتمال التحميل.',
            buttons: ['تحديث الآن', 'لاحقاً'],
            cancelId: 1
        }).then((result) => {
            if (result.response === 0) {
                autoUpdater.downloadUpdate();
            }
        });
    });

    autoUpdater.on('update-not-available', () => {
        // إظهار نافذة منبثقة تخبر المستخدم أن التطبيق محدث
        dialog.showMessageBox({
            type: 'info',
            title: 'لا يوجد تحديثات',
            message: 'أنت تستخدم أحدث إصدار من التطبيق.',
            buttons: ['حسناً']
        });
    });

    autoUpdater.on('download-progress', (progressObj) => {
        // تحديث نافذة التقدم
        let message = `السرعة: ${Math.round(progressObj.bytesPerSecond / 1024)} كيلوبايت/ثانية`;
        message += `\nتم تحميل: ${Math.round(progressObj.percent)}%`;
        message += `\n(${Math.round(progressObj.transferred / 1024)} / ${Math.round(progressObj.total / 1024)} كيلوبايت)`;
        
        // إظهار نافذة التقدم
        if (progressObj.percent === 25 || progressObj.percent === 50 || progressObj.percent === 75) {
            dialog.showMessageBox({
                type: 'info',
                title: 'تقدم التحميل',
                message: message,
                buttons: ['حسناً']
            });
        }
    });

    autoUpdater.on('update-downloaded', () => {
        // إظهار نافذة منبثقة تخبر المستخدم أن التحديث جاهز للتثبيت
        dialog.showMessageBox({
            type: 'info',
            title: 'تم تحميل التحديث',
            message: 'تم تحميل التحديث وهو جاهز للتثبيت. هل تريد إعادة تشغيل التطبيق الآن لتثبيت التحديث؟',
            buttons: ['إعادة التشغيل', 'لاحقاً'],
            cancelId: 1
        }).then((result) => {
            if (result.response === 0) {
                autoUpdater.quitAndInstall(false, true);
            }
        });
    });

    autoUpdater.on('error', (err) => {
        // إظهار نافذة منبثقة تخبر المستخدم بوجود خطأ
        dialog.showErrorBox(
            'خطأ في التحديث',
            'حدث خطأ أثناء محاولة التحديث.\n' + err.message
        );
    });
}

// إضافة قائمة التطبيق
function createMenu() {
    const template = [
        {
            label: 'التطبيق',
            submenu: [
                {
                    label: 'التحقق من التحديثات',
                    click: () => {
                        autoUpdater.checkForUpdates().catch(err => {
                            dialog.showErrorBox(
                                'خطأ في التحديث',
                                'حدث خطأ أثناء محاولة التحقق من التحديثات.\n' + err.message
                            );
                        });
                    }
                },
                { type: 'separator' },
                { role: 'quit', label: 'خروج' }
            ]
        }
    ];

    const menu = Menu.buildFromTemplate(template);
    Menu.setApplicationMenu(menu);
}

// تهيئة التحديث التلقائي عند بدء التطبيق
app.on('ready', () => {
    // تكوين التحديث التلقائي
    configureAutoUpdater();
    
    // إنشاء قائمة التطبيق
    createMenu();
    
    // التحقق من وجود تحديثات عند بدء التشغيل
    setTimeout(() => {
        autoUpdater.checkForUpdates().catch(err => {
            console.error('Error checking for updates:', err);
        });
    }, 3000); // انتظر 3 ثواني قبل التحقق من التحديثات
    
    // تنظيف الكاش عند البدء
    clearCacheOnStartup();
    
    // إنشاء النافذة
    createWindow();
    
    // تسجيل الاختصارات بعد أن يصبح التطبيق جاهزاً
    setTimeout(() => {
        registerShortcuts();
        console.log('✓ Global shortcuts registered successfully');
        
        // عرض معلومات الحفظ المسترد
        const savedMainBounds = store.get('mainWindowBounds');
        if (savedMainBounds) {
            console.log('✓ Main window restored from saved bounds:', savedMainBounds);
        }
        
        const savedTools = store.get('savedTools') || [];
        console.log(`✓ Found ${savedTools.length} saved tools`);
        
    }, 1000);
});

// دالة لتنظيف الكاش والذاكرة
async function performMemoryCleanup() {
  try {
    // تنظيف الكاش للجلسة الحالية
    if (mainWindow && !mainWindow.isDestroyed()) {
      await mainWindow.webContents.session.clearCache();
    }
    
    // تنظيف كاش جميع النوافذ المفتوحة
    for (const [path, window] of toolWindows) {
      if (window && !window.isDestroyed()) {
        await window.webContents.session.clearCache();
      }
    }
    
    // تشغيل garbage collector
    if (global.gc) {
      global.gc();
    }
    
            quietLog('Memory cleanup performed successfully');
  } catch (error) {
    console.error('Error during memory cleanup:', error);
  }
}

// بدء تنظيف دوري للذاكرة كل 5 دقائق
function startPeriodicMemoryCleanup() {
  cacheCleanupInterval = setInterval(() => {
    performMemoryCleanup();
  }, 5 * 60 * 1000); // كل 5 دقائق
}

// تنظيف الكاش عند البدء
function clearCacheOnStartup() {
  try {
    const cachePath = app.getPath('userData');
    if (fs.existsSync(cachePath)) {
      const cacheFiles = fs.readdirSync(cachePath);
      for (const file of cacheFiles) {
        if (file.includes('Cache') || file.includes('cache')) {
          const filePath = path.join(cachePath, file);
          try {
            if (fs.statSync(filePath).isDirectory()) {
              fs.rmSync(filePath, { recursive: true, force: true });
            } else {
              fs.unlinkSync(filePath);
            }
          } catch (e) {
            // تجاهل الأخطاء في حذف الملفات المقفلة
          }
        }
      }
    }
  } catch (error) {
    console.warn('Failed to clear cache on startup:', error);
  }
}

// فتح نافذة تخصيص الاختصارات
function openShortcutsConfigWindow() {
  if (shortcutsConfigWindow && !shortcutsConfigWindow.isDestroyed()) {
    shortcutsConfigWindow.focus();
    return;
  }

  shortcutsConfigWindow = new BrowserWindow({
    width: 900,
    height: 700,
    modal: false,
    parent: mainWindow,
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: '#1a1b26',
      symbolColor: '#c0caf5',
      height: 30
    },
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
      // تحسينات الأداء والحركة
      backgroundThrottling: false,
      hardwareAcceleration: false,
      webSecurity: false,
      offscreen: false,
      paintWhenInitiallyHidden: false
    },
    title: 'Shortcuts Configuration',
    show: false,
    // تحسين الحركة والأداء
    useContentSize: true,
    movable: true,
    resizable: true,
    // تقليل استهلاك الذاكرة
    enableLargerHeapSize: false
  });

  shortcutsConfigWindow.setMenu(null);
  shortcutsConfigWindow.loadFile('shortcuts-config.html');

  shortcutsConfigWindow.once('ready-to-show', () => {
    shortcutsConfigWindow.show();
  });

  shortcutsConfigWindow.on('closed', () => {
    // تنظيف الذاكرة عند إغلاق نافذة الاختصارات
    if (shortcutsConfigWindow && !shortcutsConfigWindow.isDestroyed()) {
      shortcutsConfigWindow.webContents.clearHistory();
    }
    shortcutsConfigWindow = null;
    // تشغيل garbage collection
    if (global.gc) {
      global.gc();
    }
  });

          quietLog('✓ Shortcuts configuration window opened');
}

function createWindow() {
  // محاولة استرداد إعدادات النافذة الرئيسية المحفوظة
  const savedMainBounds = store.get('mainWindowBounds');
  let mainWindowState;
  
  if (savedMainBounds) {
    console.log('Loading saved main window bounds:', savedMainBounds);
    mainWindowState = {
      x: savedMainBounds.x,
      y: savedMainBounds.y,
      width: savedMainBounds.width,
      height: savedMainBounds.height,
      manage: (window) => {
        // حفظ المقاسات عند التغيير
        const saveBounds = () => {
          if (!window.isDestroyed()) {
            const bounds = window.getBounds();
            store.set('mainWindowBounds', bounds);
            // حفظ صامت
          }
        };
        
        window.on('resize', saveBounds);
        window.on('move', saveBounds);
        window.on('maximize', saveBounds);
        window.on('unmaximize', saveBounds);
      }
    };
  } else {
    mainWindowState = windowStateKeeper({ 
      defaultWidth: 900, 
      defaultHeight: 700,
      file: 'main-window-state.json'
    });
  }

  mainWindow = new BrowserWindow({
    x: mainWindowState.x,
    y: mainWindowState.y,
    width: mainWindowState.width,
    height: mainWindowState.height,
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: '#24283b',
      symbolColor: '#c0caf5',
      height: 30
    },
    webPreferences: { 
      nodeIntegration: true, 
      contextIsolation: false, 
      webSecurity: false, 
      sandbox: false, 
      devTools: true,
      webviewTag: true,
      allowRunningInsecureContent: true,
      enableRemoteModule: true,
      // تحسينات الأداء والحركة
      backgroundThrottling: false,
      disableBlinkFeatures: 'AutomationControlled',
      v8CacheOptions: 'code',
      hardwareAcceleration: false,
      offscreen: false,
      paintWhenInitiallyHidden: false
    },
    title: 'HOPE',
    icon: path.join(__dirname, 'assets/icons/icon.png'),
    show: false
  });
  
  // Enable web security in development for better debugging
  if (process.env.NODE_ENV === 'development') {
    mainWindow.webContents.openDevTools();
  }

  if (mainWindow.webContents.isDevToolsOpened()) {
    mainWindow.webContents.closeDevTools();
  }

  mainWindowState.manage(mainWindow);
  mainWindow.loadFile('index.html');
  mainWindow.setMenu(null);

  // إظهار النافذة عند الاستعداد
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  // حفظ مقاسات النافذة الرئيسية
  const saveMainWindowBounds = () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      const bounds = mainWindow.getBounds();
      store.set('mainWindowBounds', bounds);
              // حفظ صامت
    }
  };

  // متغيرات لتتبع حالة التحريك/التغيير للنافذة الرئيسية
  let mainIsMoving = false;
  let mainIsResizing = false;
  let mainSaveTimeout = null;

  // دالة حفظ مؤجلة للنافذة الرئيسية
  const mainDeferredSave = () => {
    if (mainSaveTimeout) {
      clearTimeout(mainSaveTimeout);
    }
    mainSaveTimeout = setTimeout(() => {
      if (!mainIsMoving && !mainIsResizing) {
        // التحقق من تغيير المقاسات قبل الحفظ لتجنب الكتابة غير الضرورية
        const currentBounds = mainWindow.getBounds();
        const lastSavedBounds = store.get('mainWindowBounds');
        
        if (!lastSavedBounds || 
            currentBounds.x !== lastSavedBounds.x ||
            currentBounds.y !== lastSavedBounds.y ||
            currentBounds.width !== lastSavedBounds.width ||
            currentBounds.height !== lastSavedBounds.height) {
          saveMainWindowBounds();
        }
      }
    }, 500); // انتظار 500ms بعد توقف الحركة
  };

  // تتبع بداية ونهاية التحريك للنافذة الرئيسية
  mainWindow.on('will-move', () => {
    mainIsMoving = true;
            // حركة صامتة
  });

  mainWindow.on('moved', () => {
    mainIsMoving = false;
            // انتهاء الحركة صامت
    mainDeferredSave();
  });

  // تتبع بداية ونهاية تغيير الحجم للنافذة الرئيسية
  mainWindow.on('will-resize', () => {
    mainIsResizing = true;
  });

  mainWindow.on('resized', () => {
    mainIsResizing = false;
    mainDeferredSave();
  });

  // حفظ فوري للحالات المهمة
  mainWindow.on('maximize', saveMainWindowBounds);
  mainWindow.on('unmaximize', saveMainWindowBounds);

  mainWindow.on('closed', () => {
    // حفظ أخير للبيانات قبل الإغلاق
    console.log('Main window closing, saving final state...');
    
    // تنظيف timeout أولاً
    if (mainSaveTimeout) {
      clearTimeout(mainSaveTimeout);
      mainSaveTimeout = null;
    }
    
    // حفظ مقاسات النافذة الرئيسية
    saveMainWindowBounds();
    
    // حفظ الأدوات المفتوحة
    updateSavedToolsList();
    
    // When the main window is closed, quit the entire application.
    // The 'before-quit' event will handle saving the state.
    app.quit();
  });

  mainWindow.on('minimize', () => { minimizedState = true; });
  mainWindow.on('restore', () => { minimizedState = false; });
  mainWindow.on('focus', () => {
    console.log('Main window focused');
    
    // تسجيل استخدام النافذة الرئيسية
    const recentTools = store.get('recentTools') || [];
    const mainWindowPath = 'MainWindow';
    const existingIndex = recentTools.indexOf(mainWindowPath);
    
    if (existingIndex !== -1) {
      recentTools.splice(existingIndex, 1);
    }
    
    recentTools.unshift(mainWindowPath);
    
    if (recentTools.length > 10) {
      recentTools.pop();
    }
    
    store.set('recentTools', recentTools);
    console.log('✓ Main window marked as recently used');
    
    if (minimizedState) {
      minimizedState = false;
      Array.from(toolWindows.values()).forEach(win => {
        if (win && !win.isDestroyed()) win.reload();
      });
    }
  });

  // تأخير تحميل الأدوات المحفوظة لتسريع البدء
  app.whenReady().then(() => {
    setTimeout(() => {
    const savedTools = store.get('savedTools') || [];
      savedTools.forEach((tool, index) => {
        setTimeout(() => openToolWindow(tool.path, tool.bounds), index * 200);
      });
    }, 1500); // تأخير 1.5 ثانية قبل تحميل الأدوات
  });

  // بدء التنظيف الدوري للذاكرة
  startPeriodicMemoryCleanup();
}

function openToolWindow(toolPath, savedBounds = null) {
  if (toolWindows.has(toolPath)) {
    const existingWindow = toolWindows.get(toolPath);
    if (existingWindow && !existingWindow.isDestroyed()) {
      existingWindow.focus();
      return;
    }
  }

  const defaultBounds = { width: 1100, height: 900, x: undefined, y: undefined };
  
  // محاولة الحصول على المقاسات المحفوظة من مصادر متعددة
  let bounds = savedBounds;
  if (!bounds) {
    // البحث في المقاسات الفردية أولاً
    bounds = store.get(`toolWindow.${toolPath}.bounds`);
  }
  if (!bounds) {
    // البحث في القائمة العامة
    const savedTools = store.get('savedTools') || [];
    const savedTool = savedTools.find(tool => tool.path === toolPath);
    if (savedTool && savedTool.bounds) {
      bounds = savedTool.bounds;
    }
  }
  if (!bounds) {
    bounds = defaultBounds;
  }
  
  console.log(`Loading tool ${toolPath} with bounds:`, bounds);

  const newToolWindow = new BrowserWindow({
    ...bounds,
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: '#1a1b26', // Match the tool's background
      symbolColor: '#c0caf5',
      height: 30
    },
          webPreferences: { 
        nodeIntegration: true, 
        contextIsolation: false, 
        webSecurity: false, 
        sandbox: false, 
        devTools: true,
        enableRemoteModule: true,
        // تحسينات الأداء والحركة
        backgroundThrottling: false,
        v8CacheOptions: 'code',
        hardwareAcceleration: false,
        offscreen: false,
        paintWhenInitiallyHidden: false
      },
    title: path.basename(toolPath, '.html'), // Set title from filename
    show: false
  });

  newToolWindow.setMenu(null);
  newToolWindow.loadFile(toolPath);
  toolWindows.set(toolPath, newToolWindow);
  
  // Mostrar la ventana solo cuando esté completamente cargada
  newToolWindow.once('ready-to-show', () => {
    newToolWindow.show();
  });

  const saveBounds = () => {
    if (!newToolWindow.isDestroyed()) {
      try {
        const bounds = newToolWindow.getBounds();
        
        // حفظ المقاسات بطريقة مباشرة وبسيطة
        store.set(`toolWindow.${toolPath}.bounds`, bounds);
        
        // تحديث القائمة الفورية
        const currentSavedTools = store.get('savedTools') || [];
        const existingIndex = currentSavedTools.findIndex(tool => tool.path === toolPath);
        
        if (existingIndex >= 0) {
          // تحديث الأداة الموجودة
          currentSavedTools[existingIndex].bounds = bounds;
        } else {
          // إضافة أداة جديدة
          currentSavedTools.push({ path: toolPath, bounds: bounds });
        }
        
        store.set('savedTools', currentSavedTools);
        // حفظ صامت للأداة
        
        return true;
      } catch (error) {
        console.error(`✗ Failed to save bounds for ${toolPath}:`, error);
        return false;
      }
    }
    return false;
  };

  // متغيرات لتتبع حالة التحريك/التغيير
  let isMoving = false;
  let isResizing = false;
  let saveTimeout = null;

  // دالة حفظ مؤجلة
  const deferredSave = () => {
    if (saveTimeout) {
      clearTimeout(saveTimeout);
    }
    saveTimeout = setTimeout(() => {
      if (!isMoving && !isResizing) {
        // التحقق من تغيير المقاسات قبل الحفظ لتجنب الكتابة غير الضرورية
        const currentBounds = newToolWindow.getBounds();
        const lastSavedBounds = store.get(`toolWindow.${toolPath}.bounds`);
        
        if (!lastSavedBounds || 
            currentBounds.x !== lastSavedBounds.x ||
            currentBounds.y !== lastSavedBounds.y ||
            currentBounds.width !== lastSavedBounds.width ||
            currentBounds.height !== lastSavedBounds.height) {
          saveBounds();
        }
      }
    }, 500); // انتظار 500ms بعد توقف الحركة
  };

  // تتبع بداية ونهاية التحريك
  newToolWindow.on('will-move', () => {
    isMoving = true;
            // بداية حركة الأداة
  });

  newToolWindow.on('moved', () => {
    isMoving = false;
            // انتهاء حركة الأداة
    deferredSave();
  });

  // تتبع بداية ونهاية تغيير الحجم  
  newToolWindow.on('will-resize', () => {
    isResizing = true;
  });

  newToolWindow.on('resized', () => {
    isResizing = false;
    deferredSave();
  });

  // حفظ فوري للحالات المهمة
  newToolWindow.on('maximize', saveBounds);
  newToolWindow.on('unmaximize', saveBounds);
  newToolWindow.on('close', saveBounds); // Save bounds one last time before closing
  
  // تتبع آخر أداة مستخدمة عند التركيز
  newToolWindow.on('focus', () => {
    console.log(`Tool ${toolPath} focused`);
    
    // حفظ كآخر أداة مستخدمة
    const recentTools = store.get('recentTools') || [];
    const existingIndex = recentTools.indexOf(toolPath);
    
    if (existingIndex !== -1) {
      recentTools.splice(existingIndex, 1);
    }
    
    recentTools.unshift(toolPath);
    
    // الاحتفاظ بآخر 10 أدوات فقط
    if (recentTools.length > 10) {
      recentTools.pop();
    }
    
    store.set('recentTools', recentTools);
    console.log('✓ Updated recent tools:', recentTools);
  });
  
  // تنظيف timeout عند إغلاق النافذة
  const cleanupSaveTimeout = () => {
    if (saveTimeout) {
      clearTimeout(saveTimeout);
      saveTimeout = null;
    }
  };

  newToolWindow.on('closed', () => {
    console.log(`Tool window closed: ${toolPath}`);
    
    // تنظيف timeout أولاً
    cleanupSaveTimeout();
    
    // حفظ نهائي قبل الحذف
    if (!newToolWindow.isDestroyed()) {
      saveBounds();
    }
    
    // إزالة من القائمة
    toolWindows.delete(toolPath);
    
    // تحديث القائمة المحفوظة لإزالة النافذة المغلقة
    const currentSavedTools = store.get('savedTools') || [];
    const updatedTools = currentSavedTools.filter(tool => tool.path !== toolPath);
    store.set('savedTools', updatedTools);
    
    console.log(`Removed ${toolPath} from saved tools. Remaining: ${updatedTools.length}`);
    
    // التركيز على النافذة الرئيسية
    if (mainWindow && !mainWindow.isDestroyed() && !mainWindow.isMinimized()) {
      mainWindow.focus();
    }
  });
}

function registerShortcuts() {
  // إلغاء تسجيل جميع الاختصارات السابقة لمنع التكرار
        globalShortcut.unregisterAll();

  // تخزين مؤقت للاختصارات المسجلة
  const registered = new Set();
  
  // دالة مساعدة لتسجيل الاختصار وتجنب التعارض
  const register = (accelerator, callback) => {
    // التحقق من أن الاختصار غير مسجل بالفعل
    if (registered.has(accelerator)) {
            quietLog(`Shortcut ${accelerator} is already registered. Skipping.`);
      return;
            }
    
    // تسجيل الاختصار
    try {
      globalShortcut.register(accelerator, callback);
      registered.add(accelerator);
            quietLog(`Successfully registered shortcut: ${accelerator}`);
    } catch (e) {
      console.error(`Failed to register shortcut: ${accelerator}`, e);
            }
  };

  // 1. فتح نافذة الأدوات
  register('Control+Alt+O', () => {
    if (mainWindow) {
      mainWindow.focus();
    }
  });

  // 2. فتح أداة MOH
  register('Control+Alt+1', () => {
    openToolWindow('MOH_tools_2.html');
  });

  // 3. فتح أداة OFOQ
  register('Control+Alt+2', () => {
    openToolWindow('OFOQ_1_2.html');
  });

  // 4. فتح أداة Smart PDF
  register('Control+Alt+3', () => {
    openToolWindow('Smart_PDF_Data_Extractor_2.html');
  });

  // 5. فتح أداة UPLOAD_REQUESTED
  register('Control+Alt+4', () => {
    openToolWindow('UPLOAD_REQUESTED.html');
  });

  // 6. فتح أداة ZARA_SHIPMENT
  register('Control+Alt+5', () => {
    openToolWindow('ZARA_SHIPMENT.html');
        });

  // 7. فتح أداة AI
  register('Control+Alt+6', () => {
    openToolWindow('AI.html');
  });
  
  // 8. إغلاق النافذة الحالية
  register('Control+Alt+W', () => {
    const focusedWindow = BrowserWindow.getFocusedWindow();
    if (focusedWindow && focusedWindow !== mainWindow) {
      focusedWindow.close();
                                }
  });

  // 9. إعادة تحميل النافذة
  register('Control+Alt+R', () => {
    const focusedWindow = BrowserWindow.getFocusedWindow();
    if (focusedWindow) {
      focusedWindow.reload();
                        }
                    });

  // 10. فتح أدوات المطور
  register('Control+Alt+I', () => {
    const focusedWindow = BrowserWindow.getFocusedWindow();
    if (focusedWindow) {
      focusedWindow.webContents.toggleDevTools();
            }
        });

  // 11. حفظ (داخل نافذة الأداة)
  register('Control+Alt+S', () => {
    const focusedWindow = BrowserWindow.getFocusedWindow();
    if (focusedWindow && focusedWindow !== mainWindow) {
      focusedWindow.webContents.send('shortcut-save');
    }
  });

  // 12. فتح ملف (داخل نافذة الأداة)
  register('Control+Alt+E', () => {
    const focusedWindow = BrowserWindow.getFocusedWindow();
    if (focusedWindow && focusedWindow !== mainWindow) {
      focusedWindow.webContents.send('shortcut-open');
            }
        });

  // 13. مسح (داخل نافذة الأداة)
  register('Control+Alt+D', () => {
    const focusedWindow = BrowserWindow.getFocusedWindow();
    if (focusedWindow && focusedWindow !== mainWindow) {
      focusedWindow.webContents.send('shortcut-clear');
    }
  });

  // 14. التبديل بين النوافذ
  register('Control+Tab', () => {
    const allWindows = BrowserWindow.getAllWindows();
    const focusedWindow = BrowserWindow.getFocusedWindow();
    if (allWindows.length > 1 && focusedWindow) {
      const focusedIndex = allWindows.findIndex(w => w.id === focusedWindow.id);
      const nextIndex = (focusedIndex + 1) % allWindows.length;
      const nextWindow = allWindows[nextIndex];
      if (nextWindow) {
        nextWindow.focus();
      }
                    }
  });
  
  register('Control+Shift+Tab', () => {
    const allWindows = BrowserWindow.getAllWindows();
    const focusedWindow = BrowserWindow.getFocusedWindow();
    if (allWindows.length > 1 && focusedWindow) {
      const focusedIndex = allWindows.findIndex(w => w.id === focusedWindow.id);
      const prevIndex = (focusedIndex - 1 + allWindows.length) % allWindows.length;
      const prevWindow = allWindows[prevIndex];
      if (prevWindow) {
        prevWindow.focus();
      }
            }
        });

  // 15. فتح نافذة إعدادات الاختصارات
  register('Control+Alt+K', () => {
    openShortcutsConfigWindow();
  });

  // 16. تصغير/استعادة التطبيق
  register('Control+Alt+M', () => {
    if (mainWindow) {
      if (minimizedState) {
        mainWindow.restore();
        minimizedState = false;
        // استعادة جميع نوافذ الأدوات
        toolWindows.forEach(win => {
          if (win && !win.isDestroyed()) {
            win.restore();
          }
        });
        } else {
        // تصغير جميع النوافذ
        BrowserWindow.getAllWindows().forEach(win => {
          if (win.isMinimizable()) {
            win.minimize();
        }
        });
        minimizedState = true;
    }
    }
  });
}

function updateSavedToolsList() {
  try {
    const openTools = [];
    for (const [path, win] of toolWindows) {
      if (win && !win.isDestroyed()) {
        const bounds = win.getBounds();
        openTools.push({
    path: path,
          bounds: bounds
        });
        
        // تحديث المقاسات الفردية أيضاً
        store.set(`toolWindow.${path}.bounds`, bounds);
      }
    }
    
  store.set('savedTools', openTools);
    console.log(`✓ Updated saved tools list: ${openTools.length} tools`);
    
    return openTools;
  } catch (error) {
    console.error('✗ Error updating saved tools list:', error);
    return [];
  }
}

// ===== IPC HANDLERS =====

// Dialog handlers
ipcMain.handle('dialog:openFile', async (event, options) => {
  const lastUsedPath = store.get(LAST_USED_PATH_KEY);
  if (lastUsedPath) {
    options.defaultPath = lastUsedPath;
  }

  const { canceled, filePaths } = await dialog.showOpenDialog(BrowserWindow.fromWebContents(event.sender), options);
  if (canceled || filePaths.length === 0) {
    return null;
  }
  
  const selectedPath = filePaths[0];
  store.set(LAST_USED_PATH_KEY, path.dirname(selectedPath));
  return selectedPath;
});

ipcMain.handle('dialog:showSaveDialog', async (event, options) => {
  const lastUsedPath = store.get(LAST_USED_PATH_KEY);
  if (lastUsedPath) {
    options.defaultPath = lastUsedPath;
  }

  const { canceled, filePath } = await dialog.showSaveDialog(BrowserWindow.fromWebContents(event.sender), options);
  if (canceled) {
    return null;
  }
  
  store.set(LAST_USED_PATH_KEY, path.dirname(filePath));
  return filePath;
});

ipcMain.handle('dialog:saveFile', async (event, options) => {
  const lastUsedPath = store.get(LAST_USED_PATH_KEY);
  if (lastUsedPath) {
    options.defaultPath = lastUsedPath;
  }

  const { canceled, filePath } = await dialog.showSaveDialog(BrowserWindow.fromWebContents(event.sender), options);
  if (canceled) {
    return null;
  }

  store.set(LAST_USED_PATH_KEY, path.dirname(filePath));
  return filePath;
});

// Shell handlers
ipcMain.handle('shell:openPath', async (event, path) => {
  try {
    // Import the shell module from electron
    const { shell } = require('electron');
    await shell.openPath(path);
    return null; // Success
  } catch (error) {
    console.error('Error opening file:', error);
    return error.message;
  }
});

// File system handlers
ipcMain.handle('fs:readFile', async (event, filePath, encoding = 'utf-8') => {
  return fs.promises.readFile(filePath, encoding);
});

ipcMain.handle('fs:readBinaryFile', async (event, filePath) => {
  const buffer = await fs.promises.readFile(filePath);
  return buffer.toString('base64');
});

// Unified fs:writeFile handler that supports both string content and encoding parameter
ipcMain.handle('fs:writeFile', async (event, filePath, content, encoding = 'utf-8') => {
  await fs.promises.writeFile(filePath, content, encoding);
  return true;
});

ipcMain.handle('fs:writeBinaryFile', async (event, filePath, base64Data) => {
  try {
    let buffer;
    if (typeof base64Data === 'string') {
      // If it's a base64 string
      buffer = Buffer.from(base64Data, 'base64');
    } else if (base64Data instanceof ArrayBuffer) {
      // If it's an ArrayBuffer
      buffer = Buffer.from(base64Data);
    } else {
      // If it's already a buffer
      buffer = base64Data;
    }
    await fs.promises.writeFile(filePath, buffer);
    return true;
  } catch (error) {
    console.error('Error writing binary file:', error);
    return false;
  }
});

// Temp directory handler
ipcMain.handle('fs:getTempDir', async () => {
  const tempDir = path.join(app.getPath('temp'), 'tri-doc-excel-temp');
  
  // Ensure the directory exists
  try {
    await fs.promises.mkdir(tempDir, { recursive: true });
  } catch (error) {
    console.error('Error creating temp directory:', error);
  }
  
  return tempDir;
});

// Excel processing handlers
ipcMain.handle('excel:read-buffer', async (event, buffer) => {
    try {
        const XLSX = require('xlsx');
        const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" });
        return { success: true, data: jsonData };
    } catch (error) {
        console.error('Error reading Excel buffer:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('excel:processZaraShipment', async (event, options) => {
  try {
    console.log('Processing ZARA shipment files with Node.js');
    console.log('Reference file:', options.referenceFilePath);
    console.log('Data files:', options.dataFilePaths);
    console.log('Output file:', options.outputFilePath);
    
    // Process the files
    const result = await ExcelProcessor.processZaraShipmentFiles(
      options.referenceFilePath,
      options.dataFilePaths,
      options.outputFilePath,
      options.originalFileNames || [] // Pass original file names if available
    );
    
    console.log('Excel processing completed:', result.success ? 'Success' : 'Failed');
    return result;
  } catch (error) {
    console.error('Error in excel:processZaraShipment handler:', error);
    return {
      success: false,
      error: error.message || 'Unknown error occurred'
    };
  }
});

ipcMain.handle('excel:readFile', async (event, filePath) => {
  try {
    console.log('Reading Excel file with Node.js:', filePath);
    
    // Read the file
    const result = await ExcelProcessor.readExcelFile(filePath);
    
    console.log('Excel reading completed:', result.success ? 'Success' : 'Failed');
    return result;
  } catch (error) {
    console.error('Error in excel:readFile handler:', error);
    return {
      success: false,
      error: error.message || 'Unknown error occurred'
    };
  }
});

ipcMain.handle('excel:writeTablesFile', async (event, options) => {
  try {
    const { tables, outputFilePath } = options;
    console.log(`Writing ${tables.length} tables to Excel file: ${outputFilePath}`);
    
    // Process the files
    const result = await ExcelProcessor.writeTablesFile(tables, outputFilePath);
    
    return result;
  } catch (error) {
    console.error('Error in excel:writeTablesFile handler:', error);
    return { success: false, error: error.message };
  }
});

// Word document handlers
ipcMain.handle('word:createFromMarkdown', async (event, options) => {
  try {
    if (!docx) {
      throw new Error('docx library not available');
    }
    
    console.log('Creating Word document from Markdown');
    console.log('Output file:', options.outputFilePath);
    
    const documentContent = options.content;
    const outputFilePath = options.outputFilePath;
    
    // Create Word document
    const { Document, Packer, Paragraph, TextRun, Table, WidthType } = docx;
    const children = [];
    const lines = documentContent.split('\n');
    let currentTableMarkdown = '';
    let inTable = false;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      // Check if line is part of a table
      if (line.startsWith('|') && line.endsWith('|')) {
        if (!inTable) {
          inTable = true;
        }
        currentTableMarkdown += line + '\n';
        
        // Check if table ends (next line is not a table)
        if (i === lines.length - 1 || !lines[i + 1].startsWith('|') || !lines[i + 1].endsWith('|')) {
          inTable = false;
          
          // Process table
          const tableRows = parseMarkdownTableToDocxRows(currentTableMarkdown);
          if (tableRows.length > 0) {
            const table = new Table({
              rows: tableRows.map(cells => ({
                children: cells.map(cellText => ({
                  children: [new Paragraph({ children: [new TextRun({ text: cellText })] })]
                }))
              })),
              width: { size: 100, type: WidthType.PERCENTAGE }
            });
            children.push(table);
            children.push(new Paragraph({})); // Add empty paragraph after table
          }
          currentTableMarkdown = '';
        }
      } else if (!inTable) {
        // Process regular text
        if (line.startsWith('# ')) {
          // Heading 1
          children.push(new Paragraph({
            children: [new TextRun({ text: line.substring(2), bold: true, size: 32 })],
            spacing: { after: 200 }
          }));
        } else if (line.startsWith('## ')) {
          // Heading 2
          children.push(new Paragraph({
            children: [new TextRun({ text: line.substring(3), bold: true, size: 28 })],
            spacing: { after: 200 }
          }));
        } else if (line.startsWith('### ')) {
          // Heading 3
          children.push(new Paragraph({
            children: [new TextRun({ text: line.substring(4), bold: true, size: 24 })],
            spacing: { after: 200 }
          }));
        } else if (line.trim() === '') {
          // Empty line
          children.push(new Paragraph({}));
        } else {
          // Regular paragraph
          children.push(new Paragraph({
            children: [new TextRun({ text: line })]
          }));
        }
      }
    }
    
    const doc = new Document({
      sections: [{
        properties: {},
        children: children
      }]
    });
    
    // Generate Word document
    const buffer = await Packer.toBuffer(doc);
    await fs.promises.writeFile(outputFilePath, buffer);
    
    console.log('Word document created successfully');
    return { success: true };
  } catch (error) {
    console.error('Error creating Word document:', error);
    return {
      success: false,
      error: error.message || 'Unknown error occurred'
    };
  }
});

// Word document generation
function parseMarkdownTableToDocxRows(markdownTableText) {
  const rows = [];
  const lines = markdownTableText.split('\n').filter(line => line.trim());
  
  // Skip header separator line (e.g., |---|---|)
  const dataLines = lines.filter(line => !line.match(/^\|[-:\s|]+\|$/));
  
  dataLines.forEach(line => {
    // Remove leading/trailing | and split by |
    const cells = line.split('|').slice(1, -1).map(cell => cell.trim());
    rows.push(cells);
  });
  
  return rows;
}

ipcMain.handle('word:generateAndSave', async (event, markdownContent, defaultPath) => {
  if (!docx) {
    return { success: false, error: 'docx library is not available.' };
  }

  try {
    const { filePath } = await dialog.showSaveDialog({
      title: 'Save Word Document',
      defaultPath: defaultPath,
      filters: [{ name: 'Word Documents', extensions: ['docx'] }]
    });

    if (!filePath) {
      return { success: false, error: 'Save operation was cancelled.' };
    }

    const { Document, Packer, Paragraph, Table } = docx;

    const tables = markdownContent.split(/\n\s*\n/);
    const docxElements = [];

    tables.forEach(tableText => {
      if (tableText.trim() !== '') {
        const docxRows = parseMarkdownTableToDocxRows(tableText);
        if (docxRows.length > 0) {
          const table = new Table({
            rows: docxRows
          });
          docxElements.push(table);
          docxElements.push(new Paragraph({ text: '' })); // Spacer
        }
      }
    });

    const doc = new Document({
      sections: [{
        children: docxElements
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    await fs.promises.writeFile(filePath, buffer);

    return { success: true, filePath: filePath };
  } catch (error) {
    console.error('Error generating Word document:', error);
    return { success: false, error: error.message };
  }
});

ipcMain.handle('excel:writeJsonToExcel', async (event, ...args) => {
  try {
    let jsonData, outputFilePath;

    // Logic to handle different argument passing styles
    if (args.length === 1 && typeof args[0] === 'object' && args[0] !== null && 'jsonData' in args[0] && 'outputFilePath' in args[0]) {
      ({ jsonData, outputFilePath } = args[0]);
    } else if (args.length === 2) {
      [jsonData, outputFilePath] = args;
    } else {
      throw new Error('Invalid arguments passed to excel:writeJsonToExcel.');
    }

    // THE REAL FIX: If jsonData is a string, it's a JSON string that needs parsing.
    if (typeof jsonData === 'string') {
      jsonData = JSON.parse(jsonData);
    }

    const result = await ExcelProcessor.writeJsonToExcel(jsonData, outputFilePath);
    return result;
  } catch (error) {
    console.error(`Error in excel:writeJsonToExcel handler: ${error.message}`);
    return { success: false, error: error.message };
  }
});

// New handler to save Word documents from structured data
ipcMain.handle('save-word', async (event, { filePath, data }) => {
    if (!docx) {
        return { success: false, error: 'DOCX library is not available.' };
    }

    try {
        let contentAsString;
        if (typeof data === 'string') {
            contentAsString = data;
        } else if (typeof data === 'object' && data !== null) {
            contentAsString = JSON.stringify(data, null, 2);
        } else {
            contentAsString = String(data);
        }

        const paragraphs = contentAsString.split('\n').map(line => {
            return new docx.Paragraph({
                children: [new docx.TextRun(line)],
                style: "normal"
            });
        });

        const doc = new docx.Document({
            creator: "HOPE-Doc",
            title: "Extracted Data",
            styles: {
                paragraphStyles: [{
                    id: "normal",
                    name: "Normal",
                    basedOn: "Normal",
                    next: "Normal",
                    run: { font: "Calibri", size: 22 }, // 11pt
                    paragraph: { spacing: { after: 120 } }, // 6pt
                }],
            },
            sections: [{
                properties: {},
                children: paragraphs,
            }],
        });

        const buffer = await docx.Packer.toBuffer(doc);
        await fs.promises.writeFile(filePath, buffer);
        
        return { success: true };
    } catch (error) {
        console.error('Error creating Word file:', error);
        return { success: false, error: error.message };
    }
});

// App event handlers
app.on('ready', () => {
  // تنظيف الكاش عند البدء
  clearCacheOnStartup();
  
  // إنشاء النافذة
  createWindow();
  
  // تسجيل الاختصارات بعد أن يصبح التطبيق جاهزاً
  setTimeout(() => {
    registerShortcuts();
    console.log('✓ Global shortcuts registered successfully');
    
    // عرض معلومات الحفظ المسترد
    const savedMainBounds = store.get('mainWindowBounds');
    if (savedMainBounds) {
      console.log('✓ Main window restored from saved bounds:', savedMainBounds);
    }
    
    const savedTools = store.get('savedTools') || [];
    console.log(`✓ Found ${savedTools.length} saved tools`);
    
  }, 1000);
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// Save open tools before the application quits
app.on('before-quit', (event) => {
  console.log('=== Application Quitting - Final Save ===');
  
  try {
    // حفظ مقاسات النافذة الرئيسية أولاً
    if (mainWindow && !mainWindow.isDestroyed()) {
      const mainBounds = mainWindow.getBounds();
      store.set('mainWindowBounds', mainBounds);
      console.log('✓ Main window bounds saved:', mainBounds);
    }
    
    // حفظ مقاسات وأماكن جميع نوافذ الأدوات
    const openTools = [];
    for (const [toolPath, window] of toolWindows) {
      if (window && !window.isDestroyed()) {
        const bounds = window.getBounds();
        openTools.push({
          path: toolPath,
          bounds: bounds
        });
        
        // حفظ فردي أيضاً
        store.set(`toolWindow.${toolPath}.bounds`, bounds);
        console.log(`✓ Saved tool: ${toolPath}`, bounds);
      }
    }
    
    // حفظ القائمة الشاملة
    store.set('savedTools', openTools);
    console.log(`✓ Saved ${openTools.length} tools to storage`);
    
    // تأكيد الحفظ
    console.log('✓ All data saved successfully');
    
  } catch (error) {
    console.error('✗ Error saving application state:', error);
  }
  
  // إيقاف التنظيف الدوري
  if (cacheCleanupInterval) {
    clearInterval(cacheCleanupInterval);
  }
  
  console.log('=== Final Save Complete ===');
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
    // إعادة تسجيل الاختصارات عند إنشاء نافذة جديدة
    setTimeout(() => {
      registerShortcuts();
    }, 500);
  }
});

// Open tool windows
ipcMain.on('open-tool', (event, toolPath) => {
  openToolWindow(toolPath);
});

// App lifecycle
app.on('will-quit', () => {
    updateSavedToolsList();
    // Unregister all shortcuts.
    globalShortcut.unregisterAll();
});

// Export functions for testing
module.exports = { createWindow, openToolWindow };

ipcMain.handle('excel:convertToPdf', async (event, options) => {
  try {
    const { excelFilePaths, outputPdfPath } = options;
    if (!Array.isArray(excelFilePaths) || !outputPdfPath) {
      throw new Error('Invalid arguments: excelFilePaths (array) and outputPdfPath (string) are required.');
    }
    const result = await ExcelProcessor.convertExcelToPdf(excelFilePaths, outputPdfPath);
    return result;
  } catch (error) {
    console.error('Error in excel:convertToPdf handler:', error);
    return { success: false, error: error.message };
  }
});
