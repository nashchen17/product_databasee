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
        const excelFileInput = document.getElementById('excelFile');
        if (!excelFileInput) {
            console.error('找不到 excelFile 元素');
            return;
        }

        // Excel檔案上傳
        excelFileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                this.showUploadedFileName(file.name);
            }
            this.handleExcelUpload(e);
        });

        // 拖拽上傳功能
        const uploadArea = document.querySelector('.upload-area');
        if (uploadArea) {
            uploadArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                uploadArea.classList.add('drag-over');
            });

            uploadArea.addEventListener('dragleave', (e) => {
                e.preventDefault();
                uploadArea.classList.remove('drag-over');
            });

            uploadArea.addEventListener('drop', (e) => {
                e.preventDefault();
                uploadArea.classList.remove('drag-over');
                
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    const file = files[0];
                    if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                        file.type === 'application/vnd.ms-excel') {
                        excelFileInput.files = files;
                        this.showUploadedFileName(file.name);
                        this.handleExcelUpload({ target: { files: [file] } });
                    } else {
                        alert('請選擇Excel檔案（.xlsx 或 .xls）');
                    }
                }
            });
        }

        // 移除檔案按鈕
        const removeFileBtn = document.querySelector('.remove-file-btn');
        if (removeFileBtn) {
            removeFileBtn.addEventListener('click', () => {
                this.clearUploadedFile();
            });
        }

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
                this.currentFilter = 'all';
                const seriesSelect = document.getElementById('seriesSelect');
                if (seriesSelect) {
                    seriesSelect.value = 'all';
                }
                this.renderProductList();
            });
        }
    }

    showUploadedFileName(fileName) {
        const uploadedFileDiv = document.getElementById('uploadedFileName');
        const fileNameSpan = uploadedFileDiv.querySelector('.file-name');
        
        if (uploadedFileDiv && fileNameSpan) {
            fileNameSpan.textContent = fileName;
            uploadedFileDiv.style.display = 'flex';
        }
    }

    clearUploadedFile() {
        const excelFileInput = document.getElementById('excelFile');
        const uploadedFileDiv = document.getElementById('uploadedFileName');
        
        if (excelFileInput) {
            excelFileInput.value = '';
        }
        
        if (uploadedFileDiv) {
            uploadedFileDiv.style.display = 'none';
        }
        
        // 清除產品數據
        this.products.clear();
        this.productVersions.clear();
        this.productSeries.clear();
        this.renderProductList();
    }

    handleExcelUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 讀取第一個工作表
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                
                // 轉換為JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                console.log('Excel 數據:', jsonData);
                this.processExcelData(jsonData);
                
            } catch (error) {
                console.error('讀取Excel檔案時發生錯誤:', error);
                alert('讀取Excel檔案失敗，請確認檔案格式正確');
            }
        };
        
        reader.readAsArrayBuffer(file);
    }

    processExcelData(data) {
        if (!data || data.length === 0) {
            alert('Excel檔案為空或格式不正確');
            return;
        }

        // 清除現有數據
        this.products.clear();
        this.productVersions.clear();
        this.productSeries.clear();

        // 檢測列結構
        let seriesCol = -1, codeCol = -1, versionCol = -1, nameCol = -1, inventoryCol = -1;
        
        // 從第一行檢測標題
        if (data.length > 0) {
            const headers = data[0];
            for (let i = 0; i < headers.length; i++) {
                const header = String(headers[i] || '').trim();
                if (header.includes('系列') || header === '系列') {
                    seriesCol = i;
                } else if (header.includes('料號') || header === '料號') {
                    codeCol = i;
                } else if (header.includes('版本') || header === '版本') {
                    versionCol = i;
                } else if (header.includes('名稱') || header.includes('中文') || header === '中文名稱') {
                    nameCol = i;
                } else if (header.includes('庫存') || header === '庫存') {
                    inventoryCol = i;
                }
            }
        }

        // 如果沒有檢測到標題，使用預設位置
        if (seriesCol === -1 && codeCol === -1 && nameCol === -1) {
            seriesCol = 0;
            codeCol = 1;
            versionCol = 2;
            nameCol = 3;
            if (data[0] && data[0].length > 4) {
                inventoryCol = 4;
            }
        }

        console.log(`檢測到的列位置: 系列=${seriesCol}, 料號=${codeCol}, 版本=${versionCol}, 名稱=${nameCol}, 庫存=${inventoryCol}`);

        let addedCount = 0;
        let startRow = 1; // 從第二行開始（跳過標題行）

        // 如果第一行看起來像數據而非標題，則從第一行開始
        if (data.length > 1 && data[0]) {
            const firstRowSeries = String(data[0][seriesCol] || '').trim();
            if (firstRowSeries && (firstRowSeries.includes('系列') || /^\d+/.test(firstRowSeries))) {
                startRow = 0;
            }
        }

        for (let i = startRow; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;

            const series = String(row[seriesCol] || '').trim();
            const productCode = String(row[codeCol] || '').trim();
            const version = versionCol >= 0 ? String(row[versionCol] || '').trim() : '';
            const productName = String(row[nameCol] || '').trim();
            const inventory = inventoryCol >= 0 ? String(row[inventoryCol] || '').trim() : '';

            if (productCode && productName) {
                this.products.set(productCode, productName);
                
                if (version) {
                    this.productVersions.set(productCode, version);
                }

                if (inventory) {
                    this.productInventory.set(productCode, inventory);
                }

                // 自動檢測系列
                let detectedSeries = this.detectProductSeries(productCode, series);
                this.productSeries.set(productCode, detectedSeries);
                
                addedCount++;
            }
        }

        console.log(`從Excel讀取了 ${addedCount} 個料號`);
        
        if (addedCount > 0) {
            this.renderProductList();
            this.updateSeriesOptions();
            alert(`成功匯入 ${addedCount} 個料號`);
        } else {
            alert('未能從Excel檔案中讀取到有效的料號資料，請檢查檔案格式');
        }
    }

    detectProductSeries(productCode, series = '') {
        // 如果Excel中已經有系列資訊，優先使用
        if (series && series !== '') {
            return series;
        }

        // 根據料號前綴自動檢測系列
        if (productCode.startsWith('600')) {
            return '600系列';
        } else if (productCode.startsWith('4022')) {
            return '4022系列';
        } else if (productCode.startsWith('DZ')) {
            return 'DZ系列';
        } else if (/^[A-Z]{2,3}\d/.test(productCode)) {
            return '其他配件';
        } else {
            return '未分類';
        }
    }

    updateSeriesOptions() {
        const seriesSelect = document.getElementById('seriesSelect');
        if (!seriesSelect) return;

        // 獲取所有系列
        const allSeries = new Set();
        for (let series of this.productSeries.values()) {
            allSeries.add(series);
        }

        // 清除現有選項（保留預設選項）
        const defaultOptions = ['all', '600', '4022', 'DZ', '其他配件', 'other'];
        const optionsToRemove = [];
        for (let option of seriesSelect.options) {
            if (!defaultOptions.includes(option.value)) {
                optionsToRemove.push(option);
            }
        }
        optionsToRemove.forEach(option => option.remove());

        // 添加檢測到的新系列
        allSeries.forEach(series => {
            if (!defaultOptions.includes(series) && series !== '未分類') {
                const option = document.createElement('option');
                option.value = series;
                option.textContent = series;
                seriesSelect.appendChild(option);
            }
        });
    }

    addManualProduct() {
        const input = document.getElementById('newProductName');
        const productCode = input.value.trim();
        
        if (!productCode) {
            alert('請輸入料號');
            return;
        }
        
        if (this.products.has(productCode)) {
            alert('該料號已存在');
            return;
        }
        
        // 預設產品名稱為料號
        this.products.set(productCode, productCode);
        
        // 自動檢測系列
        const detectedSeries = this.detectProductSeries(productCode);
        this.productSeries.set(productCode, detectedSeries);
        
        input.value = '';
        this.renderProductList();
        this.updateSeriesOptions();
        
        alert(`已新增料號: ${productCode}`);
    }

    renderProductList() {
        const productListContainer = document.getElementById('productList');
        if (!productListContainer) return;

        // 清空容器
        productListContainer.innerHTML = '';

        if (this.products.size === 0) {
            productListContainer.innerHTML = '<div class="no-products">尚未匯入任何料號</div>';
            return;
        }

        // 過濾產品
        const filteredProducts = this.getFilteredProducts();
        
        if (filteredProducts.size === 0) {
            productListContainer.innerHTML = '<div class="no-products">沒有符合篩選條件的料號</div>';
            return;
        }

        // 創建主要容器
        const mainContainer = document.createElement('div');
        mainContainer.className = 'product-display-container';

        // 創建左側系列資訊區域
        const seriesInfoSection = document.createElement('div');
        seriesInfoSection.className = 'series-info-section';
        
        const currentSeriesName = this.getCurrentSeriesDisplayName();
        const productCount = filteredProducts.size;
        const totalCount = this.products.size;
        
        seriesInfoSection.innerHTML = `
            <h3>${currentSeriesName}</h3>
            <p>顯示 ${productCount} / ${totalCount} 個產品</p>
        `;

        // 創建右側產品輸入區域
        const productInputSection = document.createElement('div');
        productInputSection.className = 'product-input-section';

        // 渲染產品列表
        this.renderProductInputList(productInputSection, filteredProducts);

        // 組裝容器
        mainContainer.appendChild(seriesInfoSection);
        mainContainer.appendChild(productInputSection);
        productListContainer.appendChild(mainContainer);
    }

    getCurrentSeriesDisplayName() {
        const filterValue = this.currentFilter;
        switch (filterValue) {
            case 'all': return '全部產品';
            case '600': return '600系列產品';
            case '4022': return '4022系列產品';
            case 'DZ': return 'DZ系列產品';
            case '其他配件': return '其他配件';
            case 'other': return '未分類產品';
            default: return `${filterValue}產品`;
        }
    }

    getFilteredProducts() {
        if (this.currentFilter === 'all') {
            return this.products;
        }

        const filtered = new Map();
        for (let [productCode, productName] of this.products.entries()) {
            const series = this.productSeries.get(productCode) || '未分類';
            
            let shouldInclude = false;
            
            if (this.currentFilter === '600' && series === '600系列') {
                shouldInclude = true;
            } else if (this.currentFilter === '4022' && series === '4022系列') {
                shouldInclude = true;
            } else if (this.currentFilter === 'DZ' && series === 'DZ系列') {
                shouldInclude = true;
            } else if (this.currentFilter === '其他配件' && series === '其他配件') {
                shouldInclude = true;
            } else if (this.currentFilter === 'other' && series === '未分類') {
                shouldInclude = true;
            } else if (this.currentFilter === series) {
                shouldInclude = true;
            }
            
            if (shouldInclude) {
                filtered.set(productCode, productName);
            }
        }
        
        return filtered;
    }

    renderProductInputList(container, products) {
        const productArray = Array.from(products.entries());
        
        // 創建產品網格容器
        const gridContainer = document.createElement('div');
        gridContainer.className = 'product-input-grid';
        
        productArray.forEach(([productCode, productName]) => {
            const productVersion = this.productVersions.get(productCode) || '';
            const currentPrice = this.productPrices.get(productCode) || '';
            const currentInventory = this.productInventory.get(productCode) || '';
            
            const productItem = document.createElement('div');
            productItem.className = 'product-input-card';
            
            productItem.innerHTML = `
                <div class="product-info">
                    <div class="product-code-name">${productCode} ${productName}</div>
                    ${productVersion ? `<div class="product-version">${productVersion}</div>` : ''}
                </div>
                <div class="input-group">
                    <div class="input-field">
                        <label>單價</label>
                        <input 
                            type="number" 
                            class="price-input" 
                            placeholder="輸入單價" 
                            value="${currentPrice}"
                            data-product="${productCode}"
                            step="1"
                            min="0"
                        />
                    </div>
                    <div class="input-field">
                        <label>庫存</label>
                        <input 
                            type="number" 
                            class="inventory-input" 
                            placeholder="輸入庫存" 
                            value="${currentInventory}"
                            data-product="${productCode}"
                            step="1"
                            min="0"
                        />
                    </div>
                </div>
            `;
            
            gridContainer.appendChild(productItem);
            
            // 為輸入框添加事件監聽器
            setTimeout(() => {
                const priceInput = productItem.querySelector('.price-input');
                const inventoryInput = productItem.querySelector('.inventory-input');
                
                if (priceInput) {
                    priceInput.addEventListener('input', (e) => {
                        this.productPrices.set(productCode, e.target.value);
                    });
                }
                
                if (inventoryInput) {
                    inventoryInput.addEventListener('input', (e) => {
                        this.productInventory.set(productCode, e.target.value);
                    });
                }
            }, 0);
        });
        
        container.appendChild(gridContainer);
    }

    saveAllData() {
        const savedData = [];
        
        for (let [productCode, productName] of this.products.entries()) {
            const price = this.productPrices.get(productCode) || '';
            const inventory = this.productInventory.get(productCode) || '';
            const version = this.productVersions.get(productCode) || '';
            const series = this.productSeries.get(productCode) || '';
            
            if (price || inventory) {
                savedData.push({
                    productCode,
                    productName,
                    version,
                    series,
                    price,
                    inventory
                });
            }
        }
        
        if (savedData.length === 0) {
            alert('沒有資料需要儲存');
            return;
        }
        
        // 顯示已儲存的資料
        this.displaySavedProducts(savedData);
        
        // 可以在這裡添加將資料發送到伺服器的程式碼
        console.log('儲存的資料:', savedData);
        alert(`已儲存 ${savedData.length} 個料號的資料`);
    }

    displaySavedProducts(data) {
        const savedProductsContainer = document.getElementById('savedProducts');
        if (!savedProductsContainer) return;
        
        savedProductsContainer.innerHTML = '';
        
        if (data.length === 0) {
            savedProductsContainer.innerHTML = '<div class="no-data">尚未儲存任何資料</div>';
            return;
        }
        
        const table = document.createElement('table');
        table.className = 'saved-products-table';
        
        table.innerHTML = `
            <thead>
                <tr>
                    <th>系列</th>
                    <th>料號</th>
                    <th>版本</th>
                    <th>產品名稱</th>
                    <th>單價</th>
                    <th>庫存</th>
                </tr>
            </thead>
            <tbody>
                ${data.map(item => `
                    <tr>
                        <td>${item.series}</td>
                        <td>${item.productCode}</td>
                        <td>${item.version}</td>
                        <td>${item.productName}</td>
                        <td>${item.price ? `$${item.price}` : '-'}</td>
                        <td>${item.inventory || '-'}</td>
                    </tr>
                `).join('')}
            </tbody>
        `;
        
        savedProductsContainer.appendChild(table);
    }
}