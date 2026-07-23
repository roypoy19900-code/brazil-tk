// 订单明细处理工具 - JavaScript实现
// 基于SheetJS库进行Excel文件处理

class OrderProcessor {
    constructor() {
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.fileName = document.getElementById('fileName');
        this.fileInfo = document.getElementById('fileInfo');
        this.btnProcess = document.getElementById('btnProcess');
        this.progressContainer = document.getElementById('progressContainer');
        this.progressFill = document.getElementById('progressFill');
        this.progressText = document.getElementById('progressText');
        this.resultContainer = document.getElementById('resultContainer');
        this.resultStats = document.getElementById('resultStats');
        this.btnDownload = document.getElementById('btnDownload');
        this.errorMessage = document.getElementById('errorMessage');
        
        this.selectedFile = null;
        this.processedData = null;
        this.outputFileName = '订单明细.xlsx';
        
        this.init();
    }
    
    init() {
        // 上传区域点击事件
        this.uploadArea.addEventListener('click', () => {
            this.fileInput.click();
        });
        
        // 文件选择事件
        this.fileInput.addEventListener('change', (e) => {
            this.handleFileSelect(e.target.files[0]);
        });
        
        // 拖拽事件
        this.uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.uploadArea.classList.add('dragover');
        });
        
        this.uploadArea.addEventListener('dragleave', () => {
            this.uploadArea.classList.remove('dragover');
        });
        
        this.uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            this.uploadArea.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) {
                this.handleFileSelect(file);
            }
        });
        
        // 处理按钮点击
        this.btnProcess.addEventListener('click', () => {
            this.processFile();
        });
        
        // 下载按钮点击
        this.btnDownload.addEventListener('click', () => {
            this.downloadResult();
        });
    }
    
    handleFileSelect(file) {
        if (!file) return;
        
        // 验证文件类型
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        const validExtensions = ['.xlsx', '.xls'];
        const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
        
        if (!validExtensions.includes(fileExtension)) {
            this.showError('请选择有效的Excel文件（.xlsx或.xls格式）');
            return;
        }
        
        this.selectedFile = file;
        this.fileName.textContent = `已选择: ${file.name}`;
        this.fileInfo.classList.add('show');
        this.btnProcess.classList.add('show');
        this.hideError();
        this.hideResult();
    }
    
    async processFile() {
        if (!this.selectedFile) {
            this.showError('请先选择文件');
            return;
        }
        
        try {
            this.showProgress();
            this.updateProgress(10, '正在读取文件...');
            
            // 读取Excel文件
            const arrayBuffer = await this.readFileAsArrayBuffer(this.selectedFile);
            this.updateProgress(20, '解析Excel数据...');
            
            // 使用SheetJS解析
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // 转换为JSON（第1行为表头）
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            this.updateProgress(30, '清理和筛选数据...');
            
            // 执行数据处理步骤
            const processedData = this.processData(rawData);
            
            this.updateProgress(70, '生成订单明细表格...');
            
            // 生成A文件模板格式
            const outputWorkbook = this.generateOutputFile(processedData);
            
            this.updateProgress(90, '准备下载...');
            
            // 保存结果
            this.processedData = outputWorkbook;
            
            this.updateProgress(100, '处理完成！');
            
            // 显示结果统计
            this.showResult(processedData.length);
            
        } catch (error) {
            console.error('处理错误:', error);
            this.showError(`处理失败: ${error.message}`);
            this.hideProgress();
        }
    }
    
    readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(new Error('文件读取失败'));
            reader.readAsArrayBuffer(file);
        });
    }
    
    processData(rawData) {
        if (rawData.length < 3) {
            throw new Error('文件格式不正确，数据行数不足');
        }
        
        // 第1行是表头，第2行是说明，从第3行开始是数据
        const headers = rawData[0];
        const dataRows = rawData.slice(2); // 跳过第1行表头和第2行说明
        
        console.log('原始数据行数:', dataRows.length);
        console.log('表头:', headers);
        
        // 列映射（基于B文件的实际列位置）
        const colMap = {
            orderId: 0,           // A列 - Order ID
            orderStatus: 2,       // C列 - Order Status
            orderSubstatus: 3,    // D列 - Order Substatus
            productName: 8,       // I列 - Product Name
            quantity: 11,         // L列 - Quantity
            orderAmount: 23,      // X列 - Order Amount
            createdTime: 28,      // AC列 - Created Time
            trackingId: 36,       // AK列 - Tracking ID
            shippingProvider: 41  // AO列 - Shipping Provider Name
        };
        
        // 第一步：清理数据
        let cleanedData = dataRows.map(row => {
            const newRow = [...row];
            
            // 删除"BRL"字符，将逗号替换为点号（针对金额列）
            if (newRow[colMap.orderAmount]) {
                let amountStr = String(newRow[colMap.orderAmount]);
                amountStr = amountStr.replace(/BRL/gi, '').replace(/,/g, '.').trim();
                newRow[colMap.orderAmount] = amountStr;
            }
            
            return newRow;
        });
        
        // 第二步：筛选数据
        let filteredData = cleanedData.filter(row => {
            // 只保留Order Substatus为"已送达"的行
            const substatus = row[colMap.orderSubstatus];
            if (substatus !== '已送达') {
                return false;
            }
            
            // 剔除Order Amount为0的行
            const amountStr = String(row[colMap.orderAmount] || '0');
            const amount = parseFloat(amountStr);
            if (isNaN(amount) || amount === 0) {
                return false;
            }
            
            return true;
        });
        
        console.log('筛选后数据行数:', filteredData.length);
        
        // 第三步：格式化日期（Created Time列）
        filteredData = filteredData.map(row => {
            const newRow = [...row];
            const createdTime = row[colMap.createdTime];
            
            if (createdTime) {
                // 解析日期格式如 "07/21/2026 1:04:34 AM"
                const dateStr = String(createdTime);
                const formattedDate = this.formatDate(dateStr);
                newRow[colMap.createdTime] = formattedDate;
            }
            
            return newRow;
        });
        
        // 第四步：处理订单编号去重（模拟C文件的公式逻辑）
        const orderIdCount = {};
        const finalData = filteredData.map(row => {
            const newRow = [...row];
            const orderId = String(row[colMap.orderId]);
            
            if (!orderIdCount[orderId]) {
                orderIdCount[orderId] = 1;
                newRow[colMap.orderId] = orderId;
            } else {
                orderIdCount[orderId]++;
                newRow[colMap.orderId] = `${orderId}-${orderIdCount[orderId] - 1}`;
            }
            
            return newRow;
        });
        
        console.log('最终数据行数:', finalData.length);
        
        return finalData;
    }
    
    formatDate(dateStr) {
        // 解析格式: "07/21/2026 1:04:34 AM" -> "2026/07/21"
        try {
            const parts = dateStr.split(' ');
            const datePart = parts[0]; // "07/21/2026"
            const dateComponents = datePart.split('/');
            
            if (dateComponents.length === 3) {
                const month = dateComponents[0].padStart(2, '0');
                const day = dateComponents[1].padStart(2, '0');
                const year = dateComponents[2];
                return `${year}/${month}/${day}`;
            }
        } catch (e) {
            console.warn('日期解析失败:', dateStr);
        }
        
        return dateStr;
    }
    
    generateOutputFile(processedData) {
        // 创建输出工作簿
        const outputWB = XLSX.utils.book_new();
        
        // 准备输出数据（按照A文件模板的列顺序）
        const outputHeaders = [
            '订单编号',      // A列
            '订单日期',      // B列
            '订单币种',      // C列
            '订单金额',      // D列
            '商品名称',      // E列
            '商品数量',      // F列
            '商品单价',      // G列
            '店铺网址',      // H列
            '快递单号',      // I列
            '物流企业名称',  // J列
            '电商平台英文名称' // K列
        ];
        
        // 列映射（B文件到A文件的映射）
        const colMap = {
            orderId: 0,
            createdTime: 28,
            productName: 8,
            quantity: 11,
            orderAmount: 23,
            trackingId: 36,
            shippingProvider: 41
        };
        
        // 构建输出数据数组
        const outputData = [outputHeaders]; // 第一行是表头
        
        processedData.forEach(row => {
            // 解析订单金额
            const amountStr = String(row[colMap.orderAmount] || '0');
            const amount = parseFloat(amountStr.replace(/[^0-9.-]/g, '')) || 0;
            
            // 解析商品数量
            const quantity = parseInt(row[colMap.quantity]) || 0;
            
            // 计算商品单价
            const unitPrice = quantity > 0 ? (amount / quantity).toFixed(2) : '0.00';
            
            const outputRow = [
                row[colMap.orderId],                              // 订单编号
                row[colMap.createdTime],                          // 订单日期（已格式化）
                'BRL',                                            // 订单币种固定为BRL
                amount,                                           // 订单金额
                row[colMap.productName],                          // 商品名称
                quantity,                                         // 商品数量
                parseFloat(unitPrice),                            // 商品单价
                '',                                               // 店铺网址（空）
                row[colMap.trackingId],                           // 快递单号
                row[colMap.shippingProvider],                     // 物流企业名称
                'TK'                                              // 电商平台英文名称固定为TK
            ];
            
            outputData.push(outputRow);
        });
        
        // 创建工作表
        const outputWS = XLSX.utils.aoa_to_sheet(outputData);
        
        // 设置列宽
        outputWS['!cols'] = [
            { wch: 25 },  // 订单编号
            { wch: 15 },  // 订单日期
            { wch: 10 },  // 订单币种
            { wch: 15 },  // 订单金额
            { wch: 60 },  // 商品名称
            { wch: 12 },  // 商品数量
            { wch: 15 },  // 商品单价
            { wch: 30 },  // 店铺网址
            { wch: 20 },  // 快递单号
            { wch: 25 },  // 物流企业名称
            { wch: 15 }   // 电商平台英文名称
        ];
        
        // 添加工作表到工作簿
        XLSX.utils.book_append_sheet(outputWB, outputWS, '订单明细');
        
        return outputWB;
    }
    
    showResult(dataCount) {
        this.resultStats.innerHTML = `
            <div>✅ 成功处理 <strong>${dataCount}</strong> 条订单记录</div>
            <div>📊 已自动完成数据清理、筛选、格式化和计算</div>
            <div>💾 文件已准备就绪，点击下载按钮获取</div>
        `;
        this.resultContainer.classList.add('show');
        this.hideProgress();
    }
    
    downloadResult() {
        if (!this.processedData) {
            this.showError('没有可下载的文件');
            return;
        }
        
        try {
            // 生成Excel文件并触发下载
            XLSX.writeFile(this.processedData, this.outputFileName);
        } catch (error) {
            console.error('下载失败:', error);
            this.showError('下载失败，请重试');
        }
    }
    
    showProgress() {
        this.progressContainer.classList.add('show');
        this.btnProcess.disabled = true;
    }
    
    updateProgress(percent, text) {
        this.progressFill.style.width = percent + '%';
        this.progressText.textContent = text;
    }
    
    hideProgress() {
        this.progressContainer.classList.remove('show');
        this.btnProcess.disabled = false;
    }
    
    showError(message) {
        this.errorMessage.textContent = message;
        this.errorMessage.classList.add('show');
    }
    
    hideError() {
        this.errorMessage.classList.remove('show');
    }
    
    hideResult() {
        this.resultContainer.classList.remove('show');
    }
}

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
    new OrderProcessor();
});
