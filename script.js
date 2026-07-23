// script.js

// 全局变量存储 A 工作簿
let workbookA = null;

// 页面加载完成后立即读取内置的 A.xlsx
window.onload = async function() {
    try {
        const response = await fetch('./A.xlsx');
        if (!response.ok) throw new Error('无法加载 A.xlsx');
        
        const buffer = await response.arrayBuffer();
        workbookA = XLSX.read(buffer, { type: 'array' });
        console.log("A.xlsx 加载成功");
    } catch (error) {
        console.error("加载 A.xlsx 失败:", error);
        alert("系统初始化失败：找不到 A.xlsx 模板文件");
    }
};

document.getElementById('fileInput').addEventListener('change', handleFileSelect);

async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    // UI 状态更新
    const statusText = document.getElementById('statusText');
    const progressBar = document.querySelector('.progress-bar');
    const downloadBtn = document.getElementById('downloadBtn');
    
    statusText.innerText = "正在处理...";
    progressBar.style.width = "10%";
    downloadBtn.classList.add('hidden');

    try {
        // ==========================================
        // 步骤 1: 读取用户上传的 B 文件
        // ==========================================
        statusText.innerText = "步骤 1/4: 读取上传文件...";
        const bufferB = await file.arrayBuffer();
        const workbookB = XLSX.read(bufferB, { type: 'array' });
        const sheetNameB = workbookB.SheetNames[0];
        const dataB = XLSX.utils.sheet_to_json(workbookB.Sheets[sheetNameB]);
        progressBar.style.width = "30%";

        // ==========================================
        // 步骤 2: 读取内置的 A 文件 (从内存获取)
        // ==========================================
        statusText.innerText = "步骤 2/4: 读取内置 A 模板...";
        if (!workbookA) throw new Error("A 模板未加载");
        const sheetNameA = workbookA.SheetNames[0];
        const dataA = XLSX.utils.sheet_to_json(workbookA.Sheets[sheetNameA]);
        progressBar.style.width = "50%";

        // ==========================================
        // 步骤 3: 读取内置的 C 文件 (新增修复部分)
        // ==========================================
        statusText.innerText = "步骤 3/4: 读取内置 C 数据...";
        let dataC = [];
        try {
            // 假设你的 C 文件名为 C.xlsx，且放在同级目录
            const responseC = await fetch('./C.xlsx'); 
            if (responseC.ok) {
                const bufferC = await responseC.arrayBuffer();
                const workbookC = XLSX.read(bufferC, { type: 'array' });
                const sheetNameC = workbookC.SheetNames[0];
                dataC = XLSX.utils.sheet_to_json(workbookC.Sheets[sheetNameC]);
                console.log("C 文件读取成功，行数:", dataC.length);
            } else {
                console.warn("未找到 C.xlsx，将跳过 C 数据合并");
            }
        } catch (e) {
            console.error("读取 C 文件出错:", e);
        }
        progressBar.style.width = "70%";

        // ==========================================
        // 步骤 4: 数据合并与处理
        // ==========================================
        statusText.innerText = "步骤 4/4: 数据合并计算中...";
        
        // 为了提高性能，先将 C 数据转换为 Map (以订单编号为 Key)
        // 注意：请确保 C 文件中也有 '订单编号' 这一列，否则需修改下面的 key
        const mapC = new Map();
        dataC.forEach(row => {
            // 假设 C 文件的关联键也是 '订单编号'
            // 如果 C 文件没有表头，可能需要用 row['__EMPTY'] 等
            if(row['订单编号']) {
                mapC.set(String(row['订单编号']), row);
            }
        });

        // 遍历主数据 (通常是 B 或 A，这里假设以 B 为主表进行追加)
        // 如果你的逻辑是以 A 为主表，请改为遍历 dataA
        const finalData = dataB.map(rowB => {
            // 1. 找到对应的 A 数据 (如果需要)
            // const rowA = dataA.find(a => a['订单编号'] == rowB['订单编号']);

            // 2. 找到对应的 C 数据
            const rowC = mapC.get(String(rowB['订单编号']));

            // 3. 合并对象：基础数据 + C 的特定字段
            // 你可以手动指定要合并哪些列，例如：
            return {
                ...rowB, // 保留 B 的所有列
                
                // --- 以下是从 C 文件提取并可能重命名的列 ---
                // 假设 C 文件里有 '物流渠道' 和 '运费'，你想把它们加进来
                'C_物流渠道': rowC ? rowC['物流渠道'] || rowC['渠道'] : '', 
                'C_额外备注': rowC ? rowC['备注'] : '',
                
                // 如果 C 文件只有 version 信息，也可以加上
                'C_Version': rowC ? rowC['version'] : '' 
            };
        });

        // ==========================================
        // 生成 Excel
        // ==========================================
        const wsResult = XLSX.utils.json_to_sheet(finalData);
        const wbResult = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wbResult, wsResult, "Processed_Data");

        // 触发下载
        XLSX.writeFile(wbResult, "Processed_Result.xlsx");

        // UI 完成状态
        statusText.innerText = "处理完成！所有步骤已成功执行。";
        progressBar.style.width = "100%";
        downloadBtn.classList.remove('hidden');

    } catch (err) {
        console.error(err);
        statusText.innerText = "处理出错: " + err.message;
        progressBar.style.backgroundColor = "#ef4444";
    }
}
