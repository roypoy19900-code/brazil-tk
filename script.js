// script.js

// 全局变量存储 A 工作簿和当前工作表
let workbookA = null;
let worksheetA = null;

// 1. 页面加载时预加载 A.xlsx (模板)
window.onload = async function() {
    const statusText = document.getElementById('statusText');
    try {
        statusText.innerText = "正在初始化系统，加载内置模板...";
        const responseA = await fetch('./A.xlsx');
        if (!responseA.ok) throw new Error('无法找到 A.xlsx，请确保文件在同级目录');
        const bufferA = await responseA.arrayBuffer();
        workbookA = XLSX.read(bufferA, { type: 'array' });
        worksheetA = workbookA.Sheets[workbookA.SheetNames[0]];
        
        // 第四步：撤销工作表保护 (逻辑上移除保护属性)
        if (worksheetA['!protect']) delete worksheetA['!protect'];

        console.log("A.xlsx 加载成功");
        statusText.innerText = "✅ 系统就绪，请上传 B 文件并填写店铺 URL";
    } catch (error) {
        console.error(error);
        statusText.innerText = "❌ 初始化失败: " + error.message;
        alert("错误：无法加载内置文件 A.xlsx。请确保使用本地服务器(Live Server)运行，不要直接双击打开HTML！");
    }
};

// 2. 核心处理函数 (对应 HTML 中的 onclick="startProcessing()")
async function startProcessing() {
    const fileInput = document.getElementById('fileInput');
    const shopUrl = document.getElementById('shopUrl').value.trim();
    const statusText = document.getElementById('statusText');
    const progressBar = document.getElementById('progressBar');
    const resultArea = document.getElementById('resultArea');

    // 基础校验
    if (!shopUrl) { alert("请先填写店铺 URL！"); return; }
    if (!fileInput.files || fileInput.files.length === 0) { alert("请先上传 B 文件！"); return; }

    const file = fileInput.files[0];
    resultArea.style.display = 'none';
    progressBar.style.width = '0%';
    progressBar.style.backgroundColor = "#28a745";

    try {
        // ==========================================
        // 第一步 & 第二步 & 第三步：读取并清洗 B 文件
        // ==========================================
        statusText.innerText = "步骤 1-3/11: 正在读取并清洗 B 文件...";
        progressBar.style.width = '10%';
        
        const bufferB = await file.arrayBuffer();
        const workbookB = XLSX.read(bufferB, { type: 'array', cellDates: true });
        const wsB = workbookB.Sheets[workbookB.SheetNames[0]];
        // header:1 生成二维数组，方便按列索引操作
        let dataB = XLSX.utils.sheet_to_json(wsB, { header: 1, defval: "" }); 

        // 默认删除 B 文件第二行 (索引为1)
        if (dataB.length > 1) dataB.splice(1, 1); 

        let headerRow = dataB[0]; // 保留表头
        let cleanData = [];

        // 遍历数据行进行清洗
        for (let i = 1; i < dataB.length; i++) {
            let row = dataB[i];
            
            // 第二步：筛选 D列(Order Substatus, 索引3) 只保留 "已送达"
            if (String(row[3]).indexOf("已送达") === -1) continue;

            // 第一步：修改 X列(Order Amount, 索引23)，删除 "BRL"，"," 替换为 "."
            let cleanAmountStr = String(row[23]).replace(/BRL/g, "").replace(/,/g, ".").trim();
            let amountVal = parseFloat(cleanAmountStr);
            
            // 第二步：剔除 X列 为 0 的行
            if (isNaN(amountVal) || amountVal === 0) continue;

            // 更新清洗后的金额
            row[23] = cleanAmountStr; 

            // 第三步：修改 AC列(Created Time, 索引28)，格式化为 YYYY/MM/DD
            let rawDate = row[28]; 
            let formattedDate = "";
            if (rawDate instanceof Date) {
                formattedDate = formatDate(rawDate);
            } else {
                let d = new Date(rawDate);
                formattedDate = !isNaN(d.getTime()) ? formatDate(d) : rawDate;
            }
            row[28] = formattedDate; 

            cleanData.push(row);
        }
        dataB = [headerRow, ...cleanData];
        console.log("B 文件清洗完成，剩余有效行数:", dataB.length - 1);
        progressBar.style.width = '30%';

        // ==========================================
        // 第五步：读取 C 文件获取公式
        // ==========================================
        statusText.innerText = "步骤 5/11: 正在读取 C 文件公式...";
        const responseC = await fetch('./C.xlsx');
        const bufferC = await responseC.arrayBuffer();
        const workbookC = XLSX.read(bufferC, { type: 'array' });
        const wsC = workbookC.Sheets[workbookC.SheetNames[0]];
        
        // 假设 C 文件的 A2 单元格包含公式 (例如: ="BR-"&A1)
        let formulaStr = "";
        let cellC = wsC['A2']; 
        if (cellC && cellC.f) formulaStr = cellC.f;
        progressBar.style.width = '40%';

        // ==========================================
        // 第六步 ~ 第十步：数据映射与写入 A 文件
        // ==========================================
        statusText.innerText = "步骤 6-10/11: 正在将数据写入 A 模板...";
        let currentRowA = 2; // A文件数据写入起始索引 (对应Excel第3行，保证1、2行格式不被破坏)

        for (let i = 1; i < dataB.length; i++) {
            let rowB = dataB[i];
            
            // 第五步：使用 C 文件公式计算订单编号 -> A文件 A列 (索引0)
            let orderIdRaw = rowB[0];
            let calculatedOrderId = orderIdRaw; // 默认值
            if (formulaStr) {
                // 简单模拟 Excel 公式：将公式中的 A1 替换为实际值
                // 注意：如果是复杂公式，建议在此处直接用 JS 逻辑重写
                calculatedOrderId = formulaStr.replace(/A1/g, `"${orderIdRaw}"`).replace(/=/, "");
                // 去除首尾引号
                calculatedOrderId = calculatedOrderId.replace(/^"|"$/g, ""); 
            }
            setCellValue(currentRowA, 0, calculatedOrderId, 's');

            // 第六步：B文件 AC列(处理好的日期) -> A文件 B列 (索引1)
            setCellValue(currentRowA, 1, rowB[28], 's');

            // 第七步：B文件 X列(金额) -> A文件 D列 (索引3)
            setCellValue(currentRowA, 3, parseFloat(rowB[23]), 'n');
            
            // 第七步：A文件 C列(币种) 全部填写 BRL (索引2)
            setCellValue(currentRowA, 2, 'BRL', 's');
            
            // 第七步：B文件 L列(Quantity, 索引11) -> A文件 F列 (索引5)
            setCellValue(currentRowA, 5, parseFloat(rowB[11]), 'n');

            // 第八步：A文件 G列(单价) = D列(金额) / F列(数量) (索引6)
            let qty = parseFloat(rowB[11]) || 1; // 防止除零错误
            let price = parseFloat(rowB[23]) / qty;
            setCellValue(currentRowA, 6, price, 'n');

            // 第九步：B文件 AK列(Tracking ID, 索引36) -> A文件 I列 (索引8)
            setCellValue(currentRowA, 8, rowB[36], 's');
            
            // 第九步：B文件 AD列(Provider Name, 索引29) -> A文件 J列 (索引9)
            setCellValue(currentRowA, 9, rowB[29], 's');

            // 第十步：A文件 K列(电商平台) 填写 TK (索引10)
            setCellValue(currentRowA, 10, 'TK', 's');

            currentRowA++;
        }
        progressBar.style.width = '80%';

        // ==========================================
        // 第十一步：填写店铺 URL 到 A文件 H列
        // ==========================================
        statusText.innerText = "步骤 11/11: 正在填写店铺 URL...";
        for (let r = 2; r < currentRowA; r++) {
            setCellValue(r, 7, shopUrl, 's'); // H列索引为 7
        }

        // 更新 A 工作表的范围 (!ref)，确保包含新写入的数据
        worksheetA['!ref'] = XLSX.utils.encode_range({s: {r: 0, c: 0}, e: {r: currentRowA - 1, c: 10}});

        // ==========================================
        // 处理完成，显示下载按钮
        // ==========================================
        statusText.innerText = "✅ 处理完成！所有步骤已成功执行。";
        progressBar.style.width = '100%';
        resultArea.style.display = 'block';

    } catch (err) {
        console.error(err);
        statusText.innerText = "❌ 处理出错: " + err.message;
        progressBar.style.backgroundColor = "#ef4444";
        alert("处理过程中发生错误，请打开浏览器控制台(F12)查看详细报错。");
    }
}

// 3. 下载结果文件
function downloadResult() {
    const wbout = XLSX.write(workbookA, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), "Processed_Result.xlsx");
}

// 辅助函数：格式化日期为 YYYY/MM/DD
function formatDate(date) {
    let year = date.getFullYear();
    let month = String(date.getMonth() + 1).padStart(2, '0');
    let day = String(date.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
}

// 辅助函数：快捷设置单元格值
function setCellValue(row, col, value, type) {
    let cellRef = XLSX.utils.encode_cell({r: row, c: col});
    worksheetA[cellRef] = { t: type, v: value };
}
