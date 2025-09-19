class ProductPriceManager {
    constructor() {
        this.products = new Map(); // 存儲料號和產品資訊
        this.productVersions = new Map(); // 存儲版本資訊
        this.productSeries = new Map(); // 存儲產品系列資訊
        this.productPrices = new Map();
        this.productInventory = new Map(); // 新增庫存管理
        this.currentFilter = 'all'; // 當前篩選條件
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        // 確保DOM已加載
        const excelFileInput = document.getElementById('excelFile');
        if (!excelFileInput) {
            console.error('找不到 excelFile 元素');
            return;
        }

        // Excel檔案上傳
        excelFileInput.addEventListener('change', (e) => {
            this.handleExcelUpload(e);
        });

        // 手動新增產品
        const addProductBtn = document.getElementById('addProductBtn');
        if (addProductBtn) {
            addProductBtn.addEventListener('click', () => {
                this.addManualProduct();
            });
        }

        // 儲存所有價格和庫存
        const saveAllDataBtn = document.getElementById('saveAllData');
        if (saveAllDataBtn) {
            saveAllDataBtn.addEventListener('click', () => {
                this.saveAllData();
            });
        }

        // 產品系列篩選
        const seriesSelect = document.getElementById('seriesSelect');
        if (seriesSelect) {
            seriesSelect.addEventListener('change', (e) => {
                this.currentFilter = e.target.value;
                this.renderProductList();
            });
        }

        // 清除篩選按鈕
        const clearFilterBtn = document.getElementById('clearFilterBtn');
        if (clearFilterBtn) {
            clearFilterBtn.addEventListener('click', () => {
                const seriesSelect = document.getElementById('seriesSelect');
                if (seriesSelect) {
                    seriesSelect.value = 'all';
                    this.currentFilter = 'all';
                    this.renderProductList();
                }
            });
        }

        // Enter鍵支援
        const newProductNameInput = document.getElementById('newProductName');
        if (newProductNameInput) {
            newProductNameInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    this.addManualProduct();
                }
            });
        }

        // 添加拖放功能
        const uploadArea = document.querySelector('.upload-area');
        if (uploadArea) {
            uploadArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                uploadArea.style.borderColor = '#2980b9';
            });

            uploadArea.addEventListener('dragleave', (e) => {
                e.preventDefault();
                uploadArea.style.borderColor = '#3498db';
            });

            uploadArea.addEventListener('drop', (e) => {
                e.preventDefault();
                uploadArea.style.borderColor = '#3498db';
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    excelFileInput.files = files;
                    this.handleExcelUpload({ target: { files: files } });
                }
            });
        }
    }

    handleExcelUpload(event) {
        const file = event.target.files[0];
        if (!file) {
            this.showMessage('請選擇一個檔案', 'error');
            return;
        }

        // 檢查檔案類型
        if (!file.name.match(/\.(xlsx|xls)$/)) {
            this.showMessage('請選擇Excel檔案 (.xlsx 或 .xls)', 'error');
            return;
        }

        console.log('開始讀取檔案:', file.name);
        this.showMessage('正在讀取Excel檔案...', 'success');

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                console.log('檔案讀取完成，開始解析...');
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                console.log('工作表名稱:', workbook.SheetNames);
                
                // 讀取第一個工作表
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // 轉換為JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                console.log('解析的資料:', jsonData);
                
                if (jsonData.length === 0) {
                    this.showMessage('Excel檔案是空的', 'error');
                    return;
                }
                
                this.processExcelData(jsonData);
                this.showMessage('Excel檔案匯入成功！', 'success');
            } catch (error) {
                console.error('Excel讀取錯誤:', error);
                this.showMessage('Excel檔案讀取失敗：' + error.message, 'error');
            }
        };
        
        reader.onerror = () => {
            this.showMessage('檔案讀取失敗', 'error');
        };
        
        reader.readAsArrayBuffer(file);
    }

    processExcelData(data) {
        // 讀取系列、料號、版本和中文名稱資料
        let addedCount = 0;
        
        data.forEach((row, index) => {
            if (index === 0) return; // 跳過標題行
            
            // A欄：系列（索引0），B欄：料號（索引1），C欄：版本（索引2），D欄：中文名稱（索引3）
            const productSeries = row[0]; // A欄：系列
            const productCode = row[1]; // B欄：料號
            const productVersion = row[2]; // C欄：版本
            const productName = row[3]; // D欄：中文名稱
            
            if (productCode && typeof productCode === 'string' && productCode.trim()) {
                const trimmedCode = productCode.trim();
                const trimmedVersion = productVersion && typeof productVersion === 'string' ? productVersion.trim() : '';
                const trimmedName = productName && typeof productName === 'string' ? productName.trim() : trimmedCode;
                const trimmedSeries = productSeries && typeof productSeries === 'string' ? productSeries.trim() : '';
                
                if (!this.products.has(trimmedCode)) {
                    this.products.set(trimmedCode, trimmedName);
                    this.productVersions.set(trimmedCode, trimmedVersion);
                    
                    // 優先使用Excel中的系列資訊，如果沒有則使用自動判斷
                    let series = this.mapExcelSeriesToSystem(trimmedSeries);
                    if (!series || series === 'other') {
                        series = this.determineProductSeries(trimmedCode);
                    }
                    this.productSeries.set(trimmedCode, series);
                    
                    addedCount++;
                }
            }
        });
        
        console.log(`從Excel讀取了 ${addedCount} 個料號`);
        this.renderProductList();
        
        if (addedCount > 0) {
            this.showMessage(`成功匯入 ${addedCount} 個料號`, 'success');
        } else {
            this.showMessage('未找到有效的料號資料，請確認A欄有系列、B欄有料號、C欄有版本、D欄有中文名稱', 'error');
        }
    }

    // 將Excel中的系列名稱映射到系統系列
    mapExcelSeriesToSystem(excelSeries) {
        const series = excelSeries.toUpperCase();
        
        if (series.includes('600')) {
            return '600';
        } else if (series.includes('4022')) {
            return '4022';
        } else if (series.includes('DZ') || series.includes('DZ系列')) {
            return 'DZ';
        } else if (series.includes('配件') || series.includes('其他')) {
            return '其他配件';
        }
        
        return 'other';
    }
        
        console.log(`從Excel讀取了 ${addedCount} 個料號`);
        this.renderProductList();
        
        if (addedCount > 0) {
            this.showMessage(`成功匯入 ${addedCount} 個料號`, 'success');
        } else {
            this.showMessage('未找到有效的料號資料，請確認A欄有料號、B欄有版本、C欄有中文名稱', 'error');
        }
    }

    addManualProduct() {
        const input = document.getElementById('newProductName');
        const productCode = input.value.trim();
        
        if (!productCode) {
            this.showMessage('請輸入料號', 'error');
            return;
        }
        
        if (this.products.has(productCode)) {
            this.showMessage('料號已存在', 'error');
            return;
        }
        
        this.products.set(productCode, productCode); // 手動新增時，名稱預設為料號
        this.productVersions.set(productCode, ''); // 手動新增時，版本為空
        
        // 使用新的系列判斷方法
        const series = this.determineProductSeries(productCode);
        this.productSeries.set(productCode, series);
        
        input.value = '';
        this.renderProductList();
        this.showMessage('料號新增成功！', 'success');
    }

    // 根據料號判斷產品系列
    determineProductSeries(productCode) {
        const code = productCode.toUpperCase();
        
        // 600系列：以 "600-" 開頭的料號
        if (code.startsWith('600-')) {
            return '600';
        }
        
        // 4022系列：以 "4022." 開頭的料號
        if (code.startsWith('4022.')) {
            return '4022';
        }
        
        // DZ系列：以 "DZZN", "DZZM", "D4ZN", "D4ZM" 開頭的料號
        if (code.startsWith('DZZN') || code.startsWith('DZZM') || 
            code.startsWith('D4ZN') || code.startsWith('D4ZM')) {
            return 'DZ';
        }
        
        // 其他特殊料號（從您的清單中識別的）
        if (code.match(/^23\d{4}$/) || code.match(/^25\d{4}$/) || code.match(/^29\d{4}$/)) {
            return '其他配件';
        }
        
        return 'other';
    }

    renderProductList() {
        const container = document.getElementById('productList');
        container.innerHTML = '';
        
        // 根據篩選條件過濾產品
        const filteredProducts = new Map();
        this.products.forEach((productName, productCode) => {
            const productSeries = this.productSeries.get(productCode) || 'other';
            
            if (this.currentFilter === 'all' || productSeries === this.currentFilter) {
                filteredProducts.set(productCode, productName);
            }
        });
        
        // 顯示篩選結果統計和系列資訊
        const totalCount = this.products.size;
        const filteredCount = filteredProducts.size;
        const seriesInfo = this.getSeriesDisplayInfo(this.currentFilter);
        
        const filterHeader = document.createElement('div');
        filterHeader.className = 'filter-header';
        filterHeader.innerHTML = `
            <div class="filter-info">
                <h3 style="color: ${seriesInfo.color}; margin: 0; font-size: 18px;">
                    ${seriesInfo.title}
                </h3>
                <p style="margin: 5px 0; color: #7f8c8d;">
                    顯示 ${filteredCount} / ${totalCount} 個產品
                </p>
            </div>
        `;
        container.appendChild(filterHeader);
        
        // 如果沒有產品，顯示提示訊息
        if (filteredCount === 0) {
            const noProductsDiv = document.createElement('div');
            noProductsDiv.className = 'no-products';
            noProductsDiv.innerHTML = `
                <p style="text-align: center; color: #7f8c8d; padding: 40px;">
                    此系列暫無產品資料
                </p>
            `;
            container.appendChild(noProductsDiv);
            return;
        }
        
        // 如果是特定系列且產品數量適合表格顯示，使用表格格式
        if (this.currentFilter !== 'all' && filteredCount <= 20) {
            this.renderProductTable(container, filteredProducts, seriesInfo);
        } else {
            // 否則使用原來的網格格式
            this.renderProductGrid(container, filteredProducts);
        }
    }

    // 表格格式顯示產品
    renderProductTable(container, products, seriesInfo) {
        const tableContainer = document.createElement('div');
        tableContainer.className = 'products-table-container';
        
        const table = document.createElement('table');
        table.className = 'products-table';
        
        // 將產品轉換為陣列並按需要排列（5列4行）
        const productArray = Array.from(products.entries());
        const rows = Math.ceil(productArray.length / 4); // 4列
        
        for (let row = 0; row < Math.max(rows, 5); row++) { // 至少5行
            const tr = document.createElement('tr');
            
            for (let col = 0; col < 4; col++) { // 4列
                const index = row * 4 + col;
                const td = document.createElement('td');
                td.className = 'product-cell';
                
                if (index < productArray.length) {
                    const [productCode, productName] = productArray[index];
                    const productVersion = this.productVersions.get(productCode) || '';
                    const currentPrice = this.productPrices.get(productCode) || '';
                    const currentInventory = this.productInventory.get(productCode) || '';
                    
                    td.innerHTML = `
                        <div class="table-product-item">
                            <div class="table-product-info">
                                <div class="table-product-code">${productCode}</div>
                                ${productVersion ? `<div class="table-product-version">版本: ${productVersion}</div>` : ''}
                                <div class="table-product-name">${productName}</div>
                            </div>
                            <div class="table-input-group">
                                <input 
                                    type="number" 
                                    class="table-price-input" 
                                    placeholder="單價" 
                                    value="${currentPrice}"
                                    data-product="${productCode}"
                                    step="1"
                                    min="0"
                                />
                                <input 
                                    type="number" 
                                    class="table-inventory-input" 
                                    placeholder="庫存" 
                                    value="${currentInventory}"
                                    data-product="${productCode}"
                                    step="1"
                                    min="0"
                                />
                            </div>
                        </div>
                    `;
                    
                    // 為輸入框添加事件監聽器
                    setTimeout(() => {
                        const priceInput = td.querySelector('.table-price-input');
                        const inventoryInput = td.querySelector('.table-inventory-input');
                        
                        if (priceInput) {
                            priceInput.addEventListener('input', (e) => {
                                const price = parseInt(e.target.value);
                                if (!isNaN(price) && price >= 0) {
                                    this.productPrices.set(productCode, price);
                                } else if (e.target.value === '') {
                                    this.productPrices.delete(productCode);
                                }
                            });
                        }
                        
                        if (inventoryInput) {
                            inventoryInput.addEventListener('input', (e) => {
                                const inventory = parseInt(e.target.value);
                                if (!isNaN(inventory) && inventory >= 0) {
                                    this.productInventory.set(productCode, inventory);
                                } else if (e.target.value === '') {
                                    this.productInventory.delete(productCode);
                                }
                            });
                        }
                    }, 0);
                } else {
                    // 空儲存格
                    td.innerHTML = '<div class="empty-cell"></div>';
                }
                
                tr.appendChild(td);
            }
            
            table.appendChild(tr);
        }
        
        tableContainer.appendChild(table);
        container.appendChild(tableContainer);
    }

    // 網格格式顯示產品（原來的方式）
    renderProductGrid(container, products) {
        const productsGrid = document.createElement('div');
        productsGrid.className = 'products-grid';
        
        products.forEach((productName, productCode) => {
            const productDiv = document.createElement('div');
            productDiv.className = 'product-item';
            
            const productVersion = this.productVersions.get(productCode) || '';
            const productSeries = this.productSeries.get(productCode) || 'other';
            const currentPrice = this.productPrices.get(productCode) || '';
            const currentInventory = this.productInventory.get(productCode) || '';
            
            // 根據系列設定顏色和顯示名稱
            let seriesColor = '#7f8c8d';
            let seriesDisplayName = '未分類';
            
            if (productSeries === '600') {
                seriesColor = '#e74c3c';
                seriesDisplayName = '600系列';
            } else if (productSeries === '4022') {
                seriesColor = '#3498db';
                seriesDisplayName = '4022系列';
            } else if (productSeries === 'DZ') {
                seriesColor = '#27ae60';
                seriesDisplayName = 'DZ系列';
            } else if (productSeries === '其他配件') {
                seriesColor = '#f39c12';
                seriesDisplayName = '其他配件';
            }
            
            productDiv.innerHTML = `
                <div class="product-info">
                    <div class="product-code">${productCode}</div>
                    ${this.currentFilter !== 'all' ? '' : `<div class="product-series" style="color: ${seriesColor}">${seriesDisplayName}</div>`}
                    ${productVersion ? `<div class="product-version">版本: ${productVersion}</div>` : ''}
                    <div class="product-name">${productName}</div>
                </div>
                <div class="input-group">
                    <input 
                        type="number" 
                        class="price-input" 
                        placeholder="單價" 
                        value="${currentPrice}"
                        data-product="${productCode}"
                        step="1"
                        min="0"
                    />
                    <input 
                        type="number" 
                        class="inventory-input" 
                        placeholder="庫存" 
                        value="${currentInventory}"
                        data-product="${productCode}"
                        step="1"
                        min="0"
                    />
                </div>
            `;
            
            // 為價格輸入框添加事件監聽器
            const priceInput = productDiv.querySelector('.price-input');
            priceInput.addEventListener('input', (e) => {
                const price = parseInt(e.target.value);
                if (!isNaN(price) && price >= 0) {
                    this.productPrices.set(productCode, price);
                } else if (e.target.value === '') {
                    this.productPrices.delete(productCode);
                }
            });
            
            // 為庫存輸入框添加事件監聽器
            const inventoryInput = productDiv.querySelector('.inventory-input');
            inventoryInput.addEventListener('input', (e) => {
                const inventory = parseInt(e.target.value);
                if (!isNaN(inventory) && inventory >= 0) {
                    this.productInventory.set(productCode, inventory);
                } else if (e.target.value === '') {
                    this.productInventory.delete(productCode);
                }
            });
            
            productsGrid.appendChild(productDiv);
        });
        
        container.appendChild(productsGrid);
    }

    // 獲取系列顯示資訊
    getSeriesDisplayInfo(filter) {
        switch(filter) {
            case '600':
                return { title: '600系列產品', color: '#e74c3c' };
            case '4022':
                return { title: '4022系列產品', color: '#3498db' };
            case 'DZ':
                return { title: 'DZ系列產品', color: '#27ae60' };
            case '其他配件':
                return { title: '其他配件產品', color: '#f39c12' };
            case 'other':
                return { title: '未分類產品', color: '#7f8c8d' };
            default:
                return { title: '全部產品', color: '#2c3e50' };
        }
    }

    saveAllData() {
        const savedPriceCount = this.productPrices.size;
        const savedInventoryCount = this.productInventory.size;
        
        if (savedPriceCount === 0 && savedInventoryCount === 0) {
            this.showMessage('請至少設定一個料號的價格或庫存', 'error');
            return;
        }
        
        this.renderSavedProducts();
        this.showMessage(`成功儲存 ${savedPriceCount} 個料號的價格，${savedInventoryCount} 個料號的庫存！`, 'success');
        
        // 可以在這裡添加將資料發送到後端的邏輯
        console.log('已儲存的料號價格：', Object.fromEntries(this.productPrices));
        console.log('已儲存的料號庫存：', Object.fromEntries(this.productInventory));
    }

    renderSavedProducts() {
        const container = document.getElementById('savedProducts');
        container.innerHTML = '';
        
        // 獲取所有有資料的料號
        const allProductCodes = new Set([
            ...this.productPrices.keys(),
            ...this.productInventory.keys()
        ]);
        
        allProductCodes.forEach(productCode => {
            const productName = this.products.get(productCode) || productCode;
            const productVersion = this.productVersions.get(productCode) || '';
            const price = this.productPrices.get(productCode);
            const inventory = this.productInventory.get(productCode);
            
            const savedDiv = document.createElement('div');
            savedDiv.className = 'saved-item';
            
            let priceDisplay = price !== undefined ? `$${price.toLocaleString()}` : '未設定';
            let inventoryDisplay = inventory !== undefined ? `${inventory.toLocaleString()}` : '未設定';
            
            savedDiv.innerHTML = `
                <div class="saved-item-header">
                    <div class="saved-item-code">料號: ${productCode}</div>
                    ${productVersion ? `<div class="saved-item-version">版本: ${productVersion}</div>` : ''}
                    <div class="saved-item-name">品名: ${productName}</div>
                </div>
                <div class="saved-item-details">
                    <div class="saved-item-price">單價: ${priceDisplay}</div>
                    <div class="saved-item-inventory">庫存: ${inventoryDisplay}</div>
                </div>
            `;
            container.appendChild(savedDiv);
        });
    }

    showMessage(message, type) {
        // 移除現有的訊息
        const existingMessage = document.querySelector('.error-message, .success-message');
        if (existingMessage) {
            existingMessage.remove();
        }
        
        const messageDiv = document.createElement('div');
        messageDiv.className = type === 'error' ? 'error-message' : 'success-message';
        messageDiv.textContent = message;
        
        // 將訊息插入到第一個section之後
        const firstSection = document.querySelector('section');
        firstSection.parentNode.insertBefore(messageDiv, firstSection.nextSibling);
        
        // 3秒後自動移除訊息
        setTimeout(() => {
            messageDiv.remove();
        }, 3000);
    }

    // 匯出功能（可選）
    exportData() {
        // 獲取所有有資料的料號
        const allProductCodes = new Set([
            ...this.productPrices.keys(),
            ...this.productInventory.keys()
        ]);
        
        const data = Array.from(allProductCodes).map(code => ({
            '料號': code,
            '單價': this.productPrices.get(code) !== undefined ? `$${this.productPrices.get(code)}` : '未設定',
            '庫存': this.productInventory.get(code) !== undefined ? this.productInventory.get(code) : '未設定'
        }));
        
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "料號資料");
        XLSX.writeFile(wb, "料號價格庫存清單.xlsx");
    }
}

// 初始化應用程式
document.addEventListener('DOMContentLoaded', () => {
    new ProductPriceManager();
});
