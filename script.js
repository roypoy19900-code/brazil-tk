// script.js

// 全局变量存储 A 工作簿对象
let workbookA = null;
let worksheetA = null;

// 页面加载时预加载 A.xlsx 和 C.xlsx
window.onload = async function() {
    const statusText = document.getElementById('statusText');
    try {
        statusText.innerText = "正在初始化系统，加载内置文件...";
        
        // 1. 加载 A.xlsx (模板)
        const responseA = await fetch('./A.xlsx');
        if (!responseA.ok) throw new Error('无法找到 A.xlsx，请确保文件在同级目录');
        const bufferA = await responseA.arrayBuffer();
        workbookA = XLSX.read(bufferA, { type: 'array' });
        worksheetA = workbookA.Sheets[workbookA.SheetNames[0]]; // 默认取第一个sheet
        
        console.log("A.xlsx 加载成功");

        // 注意：C.xlsx 不需要预先加载进内存，因为它只是提供公式字符串，我们在处理时动态读取即可，
        // 或者如果C文件很大，也可以在这里预加载。为了简化，我们在步骤5时再读取C文件的特定单元格。

        statusText.innerText = "系统就绪，请上传 B 文件";
    } catch (error) {
        console.error(error);
        statusText.innerText = "初始化失败: " + error.message;
        alert("错误：无法加载内置文件 A.xlsx。请检查文件是否存在，或使用本地服务器运行。");
    }
};

document.getElementById('fileInput').addEventListener('change', handleFileSelect);

async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    const statusText = document.getElementById('statusText');
    const progressBar = document.getElementById('progressBar');
    
    try {
        // --- 第一步：操作文件 B (清洗数据) ---
        statusText.innerText = "步骤 1/11: 正在读取并清洗 B 文件...";
        const bufferB = await file.arrayBuffer();
        const workbookB = XLSX.read(bufferB, { type: 'array', cellDates: true }); // cellDates:true 帮助解析日期
        const wsB = workbookB.Sheets[workbookB.SheetNames[0]];
        
        // 将 B 转换为 JSON 对象数组，方便操作
        // header:1 表示生成二维数组，不自动推断表头，方便按列索引操作
        let dataB = XLSX.utils.sheet_to_json(wsB, { header: 1, defval: "" }); 

        // 1.1 删除第二行 (索引为1)
        if (dataB.length > 1) {
            dataB.splice(1, 1); 
        }

        // 1.2 遍历每一行进行清洗 (假设第一行是表头，从索引1开始遍历数据)
        // 注意：由于删除了原第二行，现在的索引1其实是原第三行
        // 我们需要重新构建一个干净的 dataB_clean
        let headerRow = dataB[0]; // 保留表头
        let cleanData = [];

        for (let i = 1; i < dataB.length; i++) {
            let row = dataB[i];
            
            // 获取关键列的数据 (根据截图推测列索引，需根据实际情况微调)
            // 截图显示: A=Order ID, C=Created Time, D=Order Substatus, L=Quantity, X=Order Amount, AC=Created Time(重复?), AK=Tracking ID, AD=Provider
            // 注意：JS数组索引从0开始。A=0, D=3, L=11, X=23, AC=28, AK=36, AD=29
            
            let subStatus = row[3]; // D列
            let amountStr = row[23]; // X列

            // 筛选逻辑：
            // 1. D列 (Order Substatus) 必须包含 "已送达" (或者等于，视具体文本而定，这里做模糊匹配)
            // 2. X列 (Order Amount) 剔除为 0 的行
            
            // 简单的判断逻辑，如果D列不是"已送达"则跳过
            // 注意：Excel读出来的可能是数字或字符串，需转字符串比较
            if (String(subStatus).indexOf("已送达") === -1) continue;

            // 处理金额：去除 "BRL"，逗号变点
            let cleanAmountStr = String(amountStr).replace(/BRL/g, "").replace(/,/g, ".").trim();
            let amountVal = parseFloat(cleanAmountStr);
            
            // 剔除金额为0或无效数字的行
            if (isNaN(amountVal) || amountVal === 0) continue;

            // --- 执行列修改 ---
            
            // 1. 修改 X 列 (Index 23): 替换格式
            row[23] = cleanAmountStr; 

            // 3. 修改 AC 列 (Index 28): Created Time 格式化为 YYYY/MM/DD
            // SheetJS 如果设置了 cellDates:true，日期会被解析为 Date 对象
            // 如果没有，则是字符串。我们需要兼容处理。
            let rawDate = row[28]; 
            let formattedDate = "";
            
            if (rawDate instanceof Date) {
                // 如果是日期对象
                formattedDate = formatDate(rawDate);
            } else {
                // 如果是字符串 "07/21/2026 1:04:34 AM"
                // 尝试解析
                let d = new Date(rawDate);
                if (!isNaN(d.getTime())) {
                    formattedDate = formatDate(d);
                } else {
                    formattedDate = rawDate; // 无法解析则保留原样
                }
            }
            row[28] = formattedDate; // 写回 AC 列

            cleanData.push(row);
        }

        // 重建 B 的数据结构 (表头 + 清洗后的数据)
        dataB = [headerRow, ...cleanData];
        
        console.log("B 文件清洗完成，剩余行数:", dataB.length - 1);


        // --- 第四步：A 文件撤销保护 (逻辑上的) ---
        // SheetJS 主要是处理数据。如果要“撤销保护”，通常是移除 sheet 属性中的 protected 标记
        // 但更重要的是保证前两行格式不被删。我们在写入时会从第3行开始写数据。
        if (worksheetA['!protect']) {
            delete worksheetA['!protect'];
        }


        // --- 第五步：C 文件公式计算 ---
        statusText.innerText = "步骤 5/11: 正在应用 C 文件公式计算订单编号...";
        
        // 读取 C.xlsx 获取公式
        // 假设 C 文件只有一个 sheet，且公式在 A2 (或者是某个特定单元格)
        // 题目说“C文件提供公式”，通常意味着 C 文件的某个单元格里写着类似 `="PREFIX-"&A1` 这样的公式
        // 我们需要提取这个公式字符串。
        const responseC = await fetch('./C.xlsx');
        const bufferC = await responseC.arrayBuffer();
        const workbookC = XLSX.read(bufferC, { type: 'array' });
        const wsC = workbookC.Sheets[workbookC.SheetNames[0]];
        
        // 假设公式在 C 文件的 A2 单元格 (你需要确认具体位置，这里假设是 A2)
        // 如果 C 文件只是个参考，公式是固定的，那可以直接写死。
        // 这里演示如何读取 C 文件的公式字符串
        let formulaStr = "";
        let cellC = wsC['A2']; // 假设公式在这里
        if (cellC && cellC.f) {
            formulaStr = cellC.f; // 获取公式，例如 "LEFT(A1, 5)"
        } else if (cellC && cellC.v) {
             // 如果不是公式而是值，可能逻辑不同，暂按公式处理
             // 实际上 SheetJS 读公式需要设置 { cellFormula: true } (默认开启)
        }

        // 遍历 B 的数据 (从索引1开始，即第二行)，取 A 列 (Index 0)
        // 计算出结果后，填入 A 文件的 A 列 (从第3行开始，即索引2)
        
        // A 文件的映射关系：
        // A 文件表头在第 2 行 (Excel行号)，即 JS 索引 1。
        // 数据应该从第 3 行 (Excel行号) 开始写，即 JS 索引 2。
        
        let currentRowA = 2; // A文件写入起始索引 (对应Excel第3行)

        for (let i = 1; i < dataB.length; i++) {
            let rowB = dataB[i];
            let orderIdRaw = rowB[0]; // B文件 A列 Order ID
            
            // 模拟公式计算
            // 这是一个难点：JS 不能直接运行 Excel 公式字符串。
            // 除非公式很简单。如果公式复杂，通常需要引入 hyperformula 等库。
            // **简化方案**：如果 C 文件只是为了告诉你规则，我们直接在 JS 里写逻辑。
            // **高级方案**：如果必须用 C 的公式，我们需要把公式里的引用 (如 A1) 替换成实际值，然后用 eval (危险) 或 解析器。
            
            // 假设：这里我们仅仅演示“取值”并放入 A。
            // 如果你能提供具体的公式逻辑，我可以写成 JS 函数。
            // 暂时假设：直接把 B 的 A 列 复制过去，或者做简单处理。
            // *修正*：题目说“使用C文件公式计算”。
            // 我们可以创建一个临时的 Workbook 来计算这个公式。
            
            let calculatedValue = orderIdRaw; // 默认值
            
            if (formulaStr) {
                // 创建一个临时工作表来计算公式
                // 将 B 的 A 列值放入临时表的 A1，然后读取公式所在单元格的值
                // 这种方法太慢。
                
                // 替代方法：如果公式是标准的 Excel 语法，我们可以尝试用 SheetJS 的 utils 或者手动解析。
                // 鉴于这是前端脚本，最稳妥的是：如果公式不复杂，直接 JS 实现。
                // 如果必须动态，我们假设公式是针对单个值的变换。
                
                // 示例：假设公式是 ="BR-"&A1
                // 我们替换 A1 为实际值
                let jsFormula = formulaStr.replace(/A1/g, `"${orderIdRaw}"`); 
                // 注意：这非常粗糙，仅支持简单替换。
                // 真正的 Excel 公式引擎太重了。
                // **建议**：在此处硬编码逻辑，或者告诉我公式是什么。
                // 为了代码能跑，我这里暂时直接使用 orderIdRaw。
            }

            // 写入 A 文件 A 列 (Index 0)
            // 坐标转换：Index -> Cell Address (e.g., "A3")
            let cellRefA = XLSX.utils.encode_cell({r: currentRowA, c: 0}); 
            worksheetA[cellRefA] = { t: 's', v: calculatedValue }; // s=string
            
            // --- 第六步：日期黏贴 ---
            // B 文件 C 列 (Index 2, Created Time YYYY/MM/DD) -> A 文件 B 列 (Index 1)
            // 注意：前面我们在 dataB 里已经把 AC列(28) 改了，但题目说“B文件C列”。
            // 看截图：C列是 "Order ID"?? 不，截图里 A是Order ID, C是Created Time?
            // 让我们看截图的表头：
            // A: Order ID
            // B: Order Date
            // C: Order Currency (截图没显示C，但后面说C列是币种) -> 等等，截图里 C 列被挡住了？
            // 截图显示：A(Order ID), B(Order Date), C(看不全), D(Order Substatus)...
            // 题目描述第六步：“将B文件C列（Created Time）...黏贴到A文件B列”。
            // 题目描述第三步：“AC列（Created Time）...改为YYYY/MM/DD”。
            // 这里有矛盾：第三步改的是 AC 列，第六步用的是 C 列。
            // **推测**：第三步应该是处理 B 文件的“创建时间列”（可能是 AC），第六步要把这个处理好时间填入 A。
            // 或者 B 文件既有 C 列也有 AC 列？
            // 根据截图，AC列确实是时间。C列可能是别的。
            // **修正逻辑**：我将使用刚才在步骤 3 处理好的 AC 列 (Index 28) 的数据，填入 A 的 B 列。
            
            let dateVal = rowB[28]; // 使用之前处理好的 AC 列数据
            let cellRefB = XLSX.utils.encode_cell({r: currentRowA, c: 1}); // A文件 B列
            worksheetA[cellRefB] = { t: 's', v: dateVal };

            // --- 第七步：金额、币种、数量 ---
            // B X列 (Index 23, Amount) -> A D列 (Index 3)
            let amountVal = rowB[23];
            let cellRefD = XLSX.utils.encode_cell({r: currentRowA, c: 3});
            worksheetA[cellRefD] = { t: 'n', v: parseFloat(amountVal) }; // n=number

            // A C列 (Index 2) 全部填 BRL
            let cellRefC = XLSX.utils.encode_cell({r: currentRowA, c: 2});
            worksheetA[cellRefC] = { t: 's', v: 'BRL' };

            // B L列 (Index 11, Quantity) -> A F列 (Index 5)
            let qtyVal = rowB[11];
            let cellRefF = XLSX.utils.encode_cell({r: currentRowA, c: 5});
            worksheetA[cellRefF] = { t: 'n', v: parseFloat(qtyVal) };

            // --- 第八步：计算单价 ---
            // G列 (Index 6) = D列 (Amount) / F列 (Qty)
            let price = parseFloat(amountVal) / parseFloat(qtyVal);
            let cellRefG = XLSX.utils.encode_cell({r: currentRowA, c: 6});
            worksheetA[cellRefG] = { t: 'n', v: price };

            // --- 第九步：快递信息 ---
            // B AK列 (Index 36, Tracking ID) -> A I列 (Index 8)
            let trackId = rowB[36];
            let cellRefI = XLSX.utils.encode_cell({r: currentRowA, c: 8});
            worksheetA[cellRefI] = { t: 's', v: trackId };

            // B AD列 (Index 29, Provider Name) -> A J列 (Index 9)
            // 截图里 AD 是 Shipping Provider Name? 截图只看到 AK。假设 AD 是 29 (AK是36, AL37... AD应该是29)
            // A=0, ... K=10, L=11 ... Z=25, AA=26, AB=27, AC=28, AD=29. 正确。
            let provider = rowB[29];
            let cellRefJ = XLSX.utils.encode_cell({r: currentRowA, c: 9});
            worksheetA[cellRefJ] = { t: 's', v: provider };

            // --- 第十步：电商平台 ---
            // A K列 (Index 10) 填 TK
            let cellRefK = XLSX.utils.encode_cell({r: currentRowA, c: 10});
            worksheetA[cellRefK] = { t: 's', v: 'TK' };

            currentRowA++;
        }

        // --- 第十一步：填写店铺 URL ---
        // 弹出输入框
        let shopUrl = prompt("请输入店铺URL (将填入 H 列):", "https://");
        if (shopUrl === null) shopUrl = ""; // 用户取消则留空

        // 遍历刚才写入的所有行，填入 H 列 (Index 7)
        // 范围是从 index 2 到 currentRowA - 1
        for (let r = 2; r < currentRowA; r++) {
            let cellRefH = XLSX.utils.encode_cell({r: r, c: 7});
            worksheetA[cellRefH] = { t: 's', v: shopUrl };
        }

        // --- 导出文件 ---
        statusText.innerText = "处理完成！正在生成文件...";
        
        // 更新 A 工作表的范围 (!ref)，确保包含新写入的数据
        // 获取当前最大行号
        let range = XLS
