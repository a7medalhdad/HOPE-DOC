const fs = require('fs');
const path = require('path');

// إنشاء ملف ffmpeg.dll وهمي
function createDummyFFmpeg() {
    const dummyContent = Buffer.alloc(1024, 0); // ملف فارغ 1KB
    
    // مسارات محتملة لـ ffmpeg.dll
    const possiblePaths = [
        './ffmpeg.dll',
        './resources/ffmpeg.dll',
        './node_modules/electron/dist/ffmpeg.dll',
        path.join(process.resourcesPath, 'ffmpeg.dll'),
        path.join(__dirname, 'ffmpeg.dll')
    ];
    
    possiblePaths.forEach(filePath => {
        try {
            const dir = path.dirname(filePath);
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
            }
            
            if (!fs.existsSync(filePath)) {
                fs.writeFileSync(filePath, dummyContent);
                console.log(`Created dummy ffmpeg.dll at: ${filePath}`);
            }
        } catch (error) {
            // تجاهل الأخطاء
        }
    });
}

// تشغيل الدالة
createDummyFFmpeg();

module.exports = { createDummyFFmpeg }; 