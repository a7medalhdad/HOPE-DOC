// يستخدم للتواصل مع العمليات الرئيسية
const { ipcRenderer } = require('electron');

document.addEventListener('DOMContentLoaded', () => {
    // تحميل الأدوات مباشرة بمجرد تحميل الصفحة
    initializeUI();
});

// تهيئة واجهة المستخدم
function initializeUI() {
    let toolsGrid = document.getElementById('tools-grid');
    
    // تحقق من وجود العناصر
    if (!toolsGrid) {
        console.error('Could not find tools-grid element');
        
        // إنشاء نظرة للمستخدم عند عدم وجود العناصر
        const fallbackUI = document.createElement('div');
        fallbackUI.style.cssText = 'display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; padding: 20px;';
        document.body.appendChild(fallbackUI);
        
        // استخدام fallbackUI كمرجع للعنصر الجديد
        toolsGrid = fallbackUI;
    }

    const tools = [
        { name: "PDF&ZIP", file: "MOH_tools_2.html", icon: "assets/icons/moh_tool_icon.png", description: "فك الضغط واستخراج الملفات وتقسيمها" },
        { name: "EX2XML", file: "OFOQ_1_2.html", icon: "assets/icons/ofoq_system_icon.png", description: "تحويل ملف اكسل الى نظام افق" },
        { name: "PDF2EX", file: "Smart_PDF_Data_Extractor_2.html", icon: "assets/icons/pdf_extractor_icon.png", description: "أداة لاستخراج البيانات من ملفات PDF" },
        { name: "UPLOAD REQUESTED", file: "UPLOAD_REQUESTED.html", icon: "assets/icons/upload_icon.png", description: "تحويل ملف اكسل الى نظام افق للتصاريح فقط" },
        { name: "ZARA SHIPMENT", file: "ZARA_SHIPMENT.html", icon: "assets/icons/zara_icon.png", description: "أداة خاصة بشحنات زارا" }
    ];

    // دالة فتح الأداة
    const openTool = (file) => {
        console.log(`Opening tool: ${file}`);
        // استخدام IPC لطلب فتح الأداة من العملية الرئيسية
        ipcRenderer.send('open-tool', file);
    };

    // إنشاء بطاقات الأدوات
    tools.forEach(tool => {
        const toolCard = document.createElement('div');
        toolCard.className = 'tool-card';
        
        const toolContent = document.createElement('div');
        toolContent.className = 'tool-link';
        
        toolContent.innerHTML = `
            <div class="tool-icon-container">
                <img src="${tool.icon}" alt="${tool.name}" class="tool-icon" onerror="this.src='data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iODAiIGhlaWdodD0iODAiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PGcgZmlsbD0ibm9uZSIgZmlsbC1ydWxlPSJldmVub2RkIj48cmVjdCBmaWxsPSIjN2FhMmY3IiB3aWR0aD0iODAiIGhlaWdodD0iODAiIHJ4PSI0Ii8+PHRleHQgZm9udC1mYW1pbHk9IkFyaWFsIiBmb250LXNpemU9IjM1IiBmaWxsPSIjRkZGIiB4PSI1MCUiIHk9IjUwJSIgdGV4dC1hbmNob3I9Im1pZGRsZSIgZG9taW5hbnQtYmFzZWxpbmU9Im1pZGRsZSI+JHt0b29sLm5hbWUuY2hhckF0KDApfTwvdGV4dD48L2c+PC9zdmc+'; this.classList.add('fallback-icon');">
            </div>
            <h3>${tool.name}</h3>
            <p>${tool.description}</p>
        `;
        
        // إضافة حدث النقر للبطاقة
        toolCard.appendChild(toolContent);
        toolCard.addEventListener('click', function(e) {
            e.preventDefault();
            
            // تأثير مرئي عند النقر
            this.classList.add('tool-card-active');
            
            // فتح الأداة
            openTool(tool.file);
            
            // إزالة التأثير المرئي بعد فترة
            setTimeout(() => {
                this.classList.remove('tool-card-active');
            }, 300);
        });
        
        // تأثيرات تفاعلية إضافية
        toolCard.addEventListener('mousedown', function() {
            this.style.transform = 'scale(0.95)';
        });
        
        toolCard.addEventListener('mouseup', function() {
            this.style.transform = '';
        });
        
        // إضافة البطاقة إلى الشبكة
        toolsGrid.appendChild(toolCard);
    });

    // تطبيق تأثير الظهور المتدرج للبطاقات
    document.querySelectorAll('.tool-card').forEach((card, index) => {
        card.style.animationDelay = `${index * 0.1}s`;
    });
}

// Handle update check when clicking the footer icon
document.getElementById('footer-icon').addEventListener('click', () => {
    const { ipcRenderer } = require('electron');
    ipcRenderer.send('check-for-updates');
}); 