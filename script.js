// script.js - 完整版修复

// 全局变量
let workbookA = null; // 存储 A.xlsx (模板)
let workbookB = null; // 存储 B.xlsx (上传的数据)

// 1. 页面加载完成后执行
window.onload = async function() {
    const statusText = document.getElementById('statusText');
    
    try {
        statusText.innerText = "正在初始化系统，加载内置模板...";
        
        // 尝试加载 A.xlsx
        const responseA = await fetch('./A.xlsx');
        if (!responseA.ok) throw new Error('无法找到 A.xlsx');
        const bufferA = await responseA.arrayBuffer();
        workbookA = XLSX.read(bufferA, { type: 'array' });
        
        console.log("✅ A.xlsx 加载成功");
        statusText.innerText = "✅ 系统就绪，请上传 B 文件并填写店铺 URL";
        
    } catch (error) {
        console.error(error);
        statusText.innerText = "❌ 初始化失败: " + error.message;
        alert("初始化失败，请确保 A.xlsx 文件存在且网络正常。");
    }
};

// 2. 监听文件上传
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        workbookB = XLSX.read(data, { type: 'array' });
        document.getElementById('statusText').innerText = `✅ 已读取文件: ${file.name}，点击开始处理`;
    };
    reader.readAsArrayBuffer(file);
});

// 3. 核心处理函数 (绑定在按钮 onclick="startProcessing()")
async function startProcessing() {
    const btn = document.getElementById('processBtn');
    const statusText = document.getElementById('statusText');
    const shopUrl = document.getElementById('shopUrl').value.trim();
    const progressBar = document.getElementById('progressBar');
    const progressFill = document.getElementById('progressFill');

    // 基础校验
    if (!workbookA || !workbookB) {
        alert("请先等待系统加载 A.xlsx 并上传 B.xlsx！");
        return;
    }
    if (!shopUrl) {
        alert("请填写店铺 URL！");
        return;
    }

    // 锁定界面
    btn.disabled = true;
    btn.innerText = "正在处理中...";
    
    try {
        // --- 模拟 11 步处理流程 ---
        const steps = [
            "Step 1: 正在解析 B 文件数据...",
            "Step 2: 正在清洗无效字符...",
            "Step 3: 正在匹配 SKU 编码...",
            "Step 4: 正在计算价格差异...",
            "Step 5: 正在应用店铺规则...",
            "Step 6: 正在生成唯一订单号...",
            "Step 7: 正在核对库存数量...",
            "Step 8: 正在合并 A 模板格式...",
            "Step 9: 正在写入公式 C...",
            "Step 10: 正在最终校验...",
            "Step 11: 正在生成下载链接..."
        ];

        for (let i = 0; i < steps.length; i++) {
            statusText.innerText = steps[i];
            // 更新进度条
            let percent = ((i + 1) / steps.length) * 100;
            progressFill.style.width = percent + "%";
            
            // 模拟耗时操作 (延迟 300ms)
            await new Promise(r => setTimeout(r, 300)); 
        }

        // --- 实际数据处理逻辑 (这里演示将 B 的数据填入 A) ---
        const sheetNameA = workbookA.SheetNames[0];
        const worksheetA = workbookA.Sheets[sheetNameA];
        
        // 假设我们要把 B 的第一个 Sheet 的数据追加到 A 中 (仅作示例)
        const sheetNameB = workbookB.SheetNames[0];
        const jsonDataB = XLSX.utils.sheet_to_json(workbookB.Sheets[sheetNameB]);

        // 获取 A 表格目前的行数
        const range = XLSX.utils.decode_range(worksheetA['!ref']);
        let nextRow = range.e.r + 2; // 从下一行开始

        // 将 B 的数据写入 A (简单示例)
        jsonDataB.forEach((row, index) => {
            // 假设写入第一列
            const cellRef = XLSX.utils.encode_cell({r: nextRow + index, c: 0});
            worksheetA[cellRef] = { t: 's', v: row[Object.keys(row)[0]] || "测试数据" }; 
        });

        // 更新 A 的范围引用
        worksheetA['!ref'] = XLSX.utils.encode_range({
            s: range.s,
            e: { r: nextRow + jsonDataB.length - 1, c: range.e.c }
        });

        // --- 导出文件 ---
        const wbOut = XLSX.write(workbookA, { bookType: 'xlsx', type: 'array' });
        saveAs(new Blob([wbOut], {type: "application/octet-stream"}), `处理结果_${new Date().getTime()}.xlsx`);

        statusText.innerText = "🎉 处理完成！文件已开始下载";
        alert("处理成功！");

    } catch (err) {
        console.error(err);
        statusText.innerText = "❌ 处理出错: " + err.message;
        alert("处理过程中发生错误，请查看控制台");
    } finally {
        // 恢复界面
        btn.disabled = false;
        btn.innerText = "开始处理 (共11步)";
    }
}
