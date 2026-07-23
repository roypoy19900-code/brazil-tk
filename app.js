// ================= 配置区域 =================
const APP_VERSION = "v2024.07.23"; // 在此处修改版本号
const TEMPLATE_VERSION_ROW = ["version", APP_VERSION.replace('v', '')]; // Excel中第二行的内容
// ===========================================

// 初始化页面
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('version-display').textContent = APP_VERSION;
    
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const processBtn = document.getElementById('process-btn');
    let selectedFile = null;

    // 拖拽上传逻辑
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = '#3498db';
        dropZone.style.background = '#ecf0f1';
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.style.borderColor = '#bdc3c7';
        dropZone.style.background = 'transparent';
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = '#bdc3c7';
        dropZone.style.background = 'transparent';
        if (e.dataTransfer.files.length) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFileSelect(e.target.files[0]);
        }
    });

    function handleFileSelect(file) {
        selectedFile = file;
        document.querySelector('.upload-area p').textContent = `已选择: ${file.name}`;
        processBtn.disabled = false;
        processBtn.textContent = "开始处理";
    }

    processBtn.addEventListener('click', async () => {
        if (!selectedFile) return;
        
        processBtn.disabled = true;
        processBtn.textContent = "处理中...";
        document.getElementById('status').textContent = "正在读取文件...";

        try {
            const data = await readFile(selectedFile);
            const processedData = processData(data);
            const resultWorkbook = generateOutputFile(processedData);
            downloadFile(resultWorkbook, "订单明细.xlsx");
            
            document.getElementById('status').textContent = "✅ 处理完成！文件已开始下载。";
            processBtn.textContent = "处理完成";
        } catch (error) {
            console.error(error);
            document.getElementById('status').textContent = "❌ 处理失败: " + error.message;
            processBtn.disabled = false;
            processBtn.textContent = "重试";
        }
    });
});

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            resolve(workbook);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function processData(workbook) {
    // 获取第一个 sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // 转换为 JSON (header: 1 表示生成二维数组)
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (rawData.length < 3) {
        throw new Error("文件格式不正确或数据为空");
    }

    // 1. 提取表头 (第1行)
    const headers = rawData[0];
    
    // 2. 提取数据 (从第3行开始，跳过第2行说明/版本行)
    // 注意：这里我们只取 B 文件的数据部分，A 文件的表头和版本行会在 generateOutputFile 中重新组装
    const dataRows = rawData.slice(2); 

    // 3. 数据处理逻辑 (根据你的需求：清洗、筛选、格式化等)
    const processedRows = dataRows.map(row => {
        // 示例：简单清洗，实际逻辑请根据 A/C 文件要求补充
        // 这里假设 B 文件的列顺序与 A 文件目标列有映射关系
        // 由于没有具体的 A/C 文件内容，这里保留原始数据结构，仅做基础清洗
        
        // 移除空行
        if (!row || row.length === 0) return null;
        
        // 示例：清理字符串中的 BRL 和逗号
        return row.map(cell => {
            if (typeof cell === 'string') {
                return cell.replace(/BRL/g, '').replace(/,/g, '.').trim();
            }
            return cell;
        });
    }).filter(row => row !== null);

    return {
        headers: headers,
        data: processedRows
    };
}

function generateOutputFile(processedData) {
    // 创建新的工作簿
    const wb = XLSX.utils.book_new();
    
    // 构建输出数据数组
    // 第一行：A 文件模板的表头 (这里使用 B 文件的表头作为示例，实际应替换为 A 文件固定表头)
    // 注意：如果 A 文件表头是固定的，请在这里硬编码或从其他地方加载
    const outputHeaders = processedData.headers; 
    
    const outputData = [
        outputHeaders,          // 第1行：表头
        TEMPLATE_VERSION_ROW,   // 第2行：版本号 (例如: ["version", "2024.07.23"])
        ...processedData.data   // 第3行起：处理后的数据
    ];

    const ws = XLSX.utils.aoa_to_sheet(outputData);
    
    // 自动调整列宽 (可选)
    const colWidths = outputHeaders.map(() => ({ wch: 15 }));
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, "订单明细");
    return wb;
}

function downloadFile(workbook, filename) {
    XLSX.writeFile(workbook, filename);
}
