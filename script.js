// script.js - 纯逻辑处理文件

let finalWorkbook = null;

// 页面加载检查
window.onload = function() {
    if (typeof XLSX === 'undefined') {
        console.error("SheetJS 库未加载");
    } else {
        console.log("系统就绪，等待上传...");
    }
};

async function processData() {
    const fileInput = document.getElementById('fileInput');
    const statusArea = document.getElementById('statusArea');
    const resultArea = document.getElementById('resultArea');
    const statusText = document.getElementById('statusText');
    const progressBar = document.getElementById('progressBar');
    const processBtn = document.getElementById('processBtn');

    // 基础检查
    if (!fileInput.files.length) {
        alert("⚠️ 请先选择 B 文件！");
        return;
    }

    try {
        // UI 状态更新
        processBtn.disabled = true;
        processBtn.innerText = "正在处理中...";
        statusArea.classList.remove('hidden');
        resultArea.classList.add('hidden');
        progressBar.style.width = '0%';
        statusText.innerText = "正在初始化...";

        const fileB = fileInput.files[0];

        // 1. 读取内置 A 文件
        statusText.innerText = "步骤 1/4: 读取内置模板 A.xlsx...";
        const responseA = await fetch('./A.xlsx');
        if (!responseA.ok) throw new Error("无法找到 A.xlsx，请确认文件名正确且已上传到GitHub根目录");
        const bufferA = await responseA.arrayBuffer();
        const workbookA = XLSX.read(bufferA, { type: 'array' });
        const sheetNameA = workbookA.SheetNames[0];
        const wsA = workbookA.Sheets[sheetNameA];
        
        // 将 A 转换为 JSON (header:1 表示按行读取数组，保留原始结构)
        let dataA = XLSX.utils.sheet_to_json(wsA, { header: 1 });

        // 2. 读取上传的 B 文件
        statusText.innerText = "步骤 2/4: 解析上传的数据 B...";
        const bufferB = await fileB.arrayBuffer();
        const workbookB = XLSX.read(bufferB, { type: 'array' });
        const wsB = workbookB.Sheets[workbookB.SheetNames[0]];
        const dataB = XLSX.utils.sheet_to_json(wsB); // 默认按对象读取

        // --- 开始核心逻辑处理 ---
        statusText.innerText = "步骤 3/4: 执行 11 步数据清洗与匹配...";
        progressBar.style.width = '50%';

        // 模拟 Python 的 DataFrame 操作
        
        // 1. 筛选 B 表 (删除特定行)
        // 假设 'Order Substatus' 列存在
        const filteredB = dataB.filter(row => {
            const substatus = row['Order Substatus'];
            // 排除 "运输中" 和 空值
            return substatus !== '运输中' && substatus !== undefined && substatus !== null && substatus !== '';
        });

        // 2. 计算 Processed_ID (对应 Python 的第3步)
        // 需要统计 Pure_ID 出现的次数来生成后缀
        const idCounts = {}; 
        
        // 准备一个映射字典，方便后续查找 (对应 Python 的第7-10步 set_index)
        const bMap = {}; 

        // 遍历处理后的 B 表数据
        const processedRows = filteredB.map((row, index) => {
            // 提取 Pure ID
            const rawId = String(row['OrderID'] || ''); 
            const match = rawId.match(/(\d+)/);
            const pureId = match ? match[1] : rawId;

            // 计数逻辑
            if (!idCounts[pureId]) idCounts[pureId] = 0;
            const count = idCounts[pureId];
            idCounts[pureId]++;

            // 生成 Processed ID
            const processedId = count === 0 ? pureId : `${pureId}-${count}`;

            // 存入映射表，键为 ProcessedID
            bMap[processedId] = row;

            return {
                ...row,
                Pure_ID: pureId,
                Processed_ID: processedId
            };
        });

        // 3. 填充 A 表 (对应 Python 的 loc 操作)
        // 假设 A 表的前两行是标题，数据从第3行(index 2)开始
        // 注意：这里我们需要动态确定 A 表的表头索引，或者硬编码列名映射
        
        // 获取 A 表的第一行作为表头（假设第一行是表头）
        const headersA = dataA[0]; 
        
        // 找到各列在 A 表中的索引
        const colIndex = {
            OrderID: headersA.indexOf('OrderID'),
            OrderDate: headersA.indexOf('OrderDate'),
            Currency: headersA.indexOf('Currency'),
            Warehouse: headersA.indexOf('Warehouse'),
            OrderAmount: headersA.indexOf('OrderAmount'),
            ProductName: headersA.indexOf('ProductName'),
            Quantity: headersA.indexOf('Quantity'),
            TrackingID: headersA.indexOf('TrackingID'),
            ShippingProvider: headersA.indexOf('ShippingProvider')
        };

        // 从 A 表的第3行 (index 2) 开始写入
        let aRowIndex = 2; 

        processedRows.forEach(bRow => {
            // 确保 A 表有足够的行数，如果不够就推入空行
            while (aRowIndex >= dataA.length) {
                // 创建一个空行，长度与表头一致
                dataA.push(new Array(headersA.length).fill(''));
            }

            const currentRowA = dataA[aRowIndex];

            // --- 第3步: 写入 OrderID ---
            if (colIndex.OrderID !== -1) currentRowA[colIndex.OrderID] = bRow.Processed_ID;

            // --- 第4步: 日期转换 ---
            if (colIndex.OrderDate !== -1 && bRow.ShipDate) {
                // 简单处理日期，Excel日期可能是数字或字符串
                let dateStr = bRow.ShipDate;
                // 如果是 Excel 序列号 (数字)
                if (typeof dateStr === 'number') {
                    const date = XLSX.SSF.parse_date_code(dateStr);
                    dateStr = `${date.y}/${String(date.m).padStart(2,'0')}/${String(date.d).padStart(2,'0')}`;
                } 
                // 如果是 JS Date 对象或其他格式，这里简化处理，假设输入是标准格式或需手动调整
                // 实际生产中可能需要更复杂的日期解析库，这里尝试直接格式化
                else {
                     // 尝试直接转为 Date 对象格式化
                     const d = new Date(dateStr);
                     if (!isNaN(d.getTime())) {
                         dateStr = `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
                     }
                }
                currentRowA[colIndex.OrderDate] = dateStr;
            }

            // --- 第5步: 固定字段 ---
            if (colIndex.Currency !== -1) currentRowA[colIndex.Currency] = 'BRL';
            if (colIndex.Warehouse !== -1) currentRowA[colIndex.Warehouse] = 'TK';

            // --- 第6步: 金额清洗 ---
            if (colIndex.OrderAmount !== -1 && bRow['Order Amount']) {
                let amount = String(bRow['Order Amount']).replace('BRL', '').replace(',', '.').trim();
                currentRowA[colIndex.OrderAmount] = parseFloat(amount) || 0;
            }

            // --- 第7-10步: 匹配复制数据 (直接从当前 bRow 获取，因为我们已经建立了映射) ---
            // 注意：Python 代码里是用 Processed_ID 去反查，但在这里我们是在遍历 processedRows，所以直接用 bRow 即可
            
            if (colIndex.ProductName !== -1) currentRowA[colIndex.ProductName] = bRow['Product Name'] || '';
            if (colIndex.Quantity !== -1) currentRowA[colIndex.Quantity] = bRow['Quantity'] || 0;
            if (colIndex.TrackingID !== -1) currentRowA[colIndex.TrackingID] = bRow['Tracking ID'] || '';
            if (colIndex.ShippingProvider !== -1) currentRowA[colIndex.ShippingProvider] = bRow['Shipping Provider Name'] || '';

            aRowIndex++;
        });

        // --- 第9步: 计算单价 (UnitPrice) ---
        // 需要先找到 UnitPrice 列的索引
        const unitPriceIdx = headersA.indexOf('UnitPrice');
        const amountIdx = colIndex.OrderAmount;
        const qtyIdx = colIndex.Quantity;

        if (unitPriceIdx !== -1 && amountIdx !== -1 && qtyIdx !== -1) {
            // 从数据行开始遍历 (跳过表头)
            for (let i = 1; i < dataA.length; i++) {
                const row = dataA[i];
                const amount = parseFloat(row[amountIdx]);
                const qty = parseFloat(row[qtyIdx]);
                
                // 只有当这一行是我们刚才写入的数据时才计算（或者根据业务逻辑全量计算）
                // 这里假设只计算非空行
                if (!isNaN(amount) && !isNaN(qty) && qty !== 0) {
                    row[unitPriceIdx] = amount / qty;
                }
            }
        }

        // --- 第11步: 删除 0 金额行 ---
        // 过滤掉 OrderAmount 为 0 的行 (保留表头)
        const finalData = [headersA]; // 先放入表头
        for (let i = 1; i < dataA.length; i++) {
            const row = dataA[i];
            const amount = parseFloat(row[amountIdx]);
            // 如果金额不是0，或者是空行（可能是模板自带的），保留
            // 这里的逻辑稍微复杂，假设我们要删除的是刚才填入且金额为0的行
            // 简单处理：只要 Amount 列不为 0 就保留
            if (amount !== 0) {
                finalData.push(row);
            }
        }

        // --- 生成结果 ---
        statusText.innerText = "步骤 4/4: 生成 Excel 文件...";
        progressBar.style.width = '90%';

        const newWs = XLSX.utils.aoa_to_sheet(finalData);
        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, newWs, "Processed_Data");

        finalWorkbook = newWb;

        // 完成
        progressBar.style.width = '100%';
        statusText.innerText = "✅ 处理完成！";
        processBtn.innerText = "重新处理";
        processBtn.disabled = false;
        
        resultArea.classList.remove('hidden');

    } catch (error) {
        console.error(error);
        alert("❌ 发生错误: " + error.message);
        statusText.innerText = "处理失败";
        processBtn.disabled = false;
        processBtn.innerText = "开始处理";
    }
}

function downloadFile() {
    if (finalWorkbook) {
        XLSX.writeFile(finalWorkbook, "A_final_processed.xlsx");
    }
}
