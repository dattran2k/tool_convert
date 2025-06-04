class AutomationSuite {
    constructor() {
        this.weatherData = null;
        this.templateData = null;
        this.flattenedData = null;
        this.convertedData = null;
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        // Auto file inputs
        document.getElementById('autoWeatherInput').addEventListener('change', (e) => this.handleWeatherFile(e.target.files[0]));
        document.getElementById('autoTemplateInput').addEventListener('change', (e) => this.handleTemplateFile(e.target.files[0]));
    }

    async handleWeatherFile(file) {
        if (!file) return;

        this.showAutoStatus('info', 'Đang đọc Weather Functional Spec...', true);

        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, {
                cellStyles: true,
                cellFormulas: true,
                cellDates: true,
                cellNF: true,
                sheetStubs: true
            });

            // Tìm sheet Functional_Spec
            const functionalSpecSheet = workbook.SheetNames.find(name => 
                name.toLowerCase().includes('functional') || 
                name.toLowerCase().includes('spec') ||
                name.includes('仕様')
            );

            if (!functionalSpecSheet) {
                throw new Error('Không tìm thấy sheet Functional_Spec');
            }

            const worksheet = workbook.Sheets[functionalSpecSheet];
            const range = XLSX.utils.decode_range(worksheet['!ref']);

            // Đọc tất cả dữ liệu
            let allData = [];
            for (let R = range.s.r; R <= range.e.r; ++R) {
                let rowData = [];
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({c: C, r: R});
                    const cell = worksheet[cellAddress];
                    rowData.push(cell ? (cell.v || '') : '');
                }
                allData.push(rowData);
            }

            // Tìm header row
            let headerRowIndex = -1;
            for (let i = 0; i < allData.length; i++) {
                if (allData[i][0] === 'No.' && allData[i][1] === 'Chapter') {
                    headerRowIndex = i;
                    break;
                }
            }

            const headers = allData[headerRowIndex];
            const dataRows = allData.slice(headerRowIndex + 1);
            const weatherRows = dataRows.filter(row => row[0] && row[0].startsWith('WEA_'));

            this.weatherData = {
                headers: headers,
                rows: weatherRows
            };

            document.getElementById('autoWeatherUpload').classList.add('loaded');
            document.getElementById('autoWeatherInfo').style.display = 'block';
            document.getElementById('autoWeatherInfo').textContent = 
                `✅ Đã load ${weatherRows.length} dòng WEA`;

            this.showAutoStatus('success', `Đã đọc ${weatherRows.length} weather requirements`);
            this.checkAutoReady();

        } catch (error) {
            this.showAutoStatus('error', `Lỗi đọc weather file: ${error.message}`);
        }
    }

    async handleTemplateFile(file) {
        if (!file) return;

        this.showAutoStatus('info', 'Đang đọc Requirements Template...', true);

        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, {
                cellStyles: true,
                cellFormulas: true,
                cellDates: true,
                cellNF: true,
                sheetStubs: true
            });

            // Tìm sheet 要件情報
            const requirementsSheet = workbook.SheetNames.find(name => 
                name.includes('要件情報') || name.includes('要件')
            );

            if (!requirementsSheet) {
                throw new Error('Không tìm thấy sheet 要件情報');
            }

            this.templateData = {
                workbook: workbook,
                sheetName: requirementsSheet,
                modelColumnMap: this.buildTemplateModelMap(workbook.Sheets[requirementsSheet])
            };

            document.getElementById('autoTemplateUpload').classList.add('loaded');
            document.getElementById('autoTemplateInfo').style.display = 'block';
            document.getElementById('autoTemplateInfo').textContent = 
                `✅ Template loaded: Sheet "${requirementsSheet}"`;

            this.showAutoStatus('success', 'Đã đọc template thành công');
            this.checkAutoReady();

        } catch (error) {
            this.showAutoStatus('error', `Lỗi đọc template file: ${error.message}`);
        }
    }

    buildTemplateModelMap(worksheet) {
        const modelMap = {};
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({c: C, r: 1});
            const cell = worksheet[cellAddress];
            if (cell && cell.v && typeof cell.v === 'string') {
                const value = cell.v.toString();
                const modelMatch = value.match(/([EG]B\d{4}[VU])/); 
                if (modelMatch) {
                    modelMap[modelMatch[1]] = C;
                }
            }
        }
        
        return modelMap;
    }

    checkAutoReady() {
        const autoProcessBtn = document.getElementById('autoProcessBtn');
        if (this.weatherData && this.templateData) {
            autoProcessBtn.disabled = false;
        } else {
            autoProcessBtn.disabled = true;
        }
    }

    async runAutomation() {
        if (!this.weatherData || !this.templateData) return;

        try {
            this.setProgress(0);
            document.getElementById('autoProgress').style.display = 'block';

            // Step 1: Flatten
            this.showAutoStatus('info', 'Bước 1/3: Đang flatten cấu trúc tree...', true);
            this.setProgress(20);
            await this.sleep(500);

            this.flattenedData = this.performFlatten();
            this.setProgress(40);

            // Step 2: Convert 
            this.showAutoStatus('info', 'Bước 2/3: Đang convert sang requirements format...', true);
            await this.sleep(500);

            this.convertedData = this.performConvert();
            this.setProgress(80);

            // Step 3: Export
            this.showAutoStatus('info', 'Bước 3/3: Đang export file Excel...', true);
            await this.sleep(500);

            this.performExport();
            this.setProgress(100);

            this.showAutoStatus('success', '🎉 Hoàn thành! File đã được tải về thành công.');

        } catch (error) {
            this.showAutoStatus('error', `Lỗi trong quá trình xử lý: ${error.message}`);
        }
    }

    performFlatten() {
        const flattened = [];
        let currentChapter = '';
        let currentSection = '';
        let currentSubsection = '';

        for (const row of this.weatherData.rows) {
            const [no, chapter, section, subsection, spec, ...rest] = row;
            
            if (chapter && chapter.trim() !== '') {
                currentChapter = chapter.trim();
                currentSection = '';
                currentSubsection = '';
            }
            if (section && section.trim() !== '') {
                currentSection = section.trim();
                currentSubsection = '';
            }
            if (subsection && subsection.trim() !== '') {
                currentSubsection = subsection.trim();
            }

            const flatRow = [
                no,
                currentChapter,
                currentSection,
                currentSubsection,
                spec,
                ...rest
            ];

            flattened.push(flatRow);
        }

        return flattened;
    }

    performConvert() {
        const converted = [];
        let rowIndex = 1;

        // Lọc chỉ các dòng có Functional Specification
        const specRows = this.flattenedData.filter(row => row[4] && row[4].trim() !== '');

        for (const row of specRows) {
            const [no, chapter, section, subsection, spec, link, tag, ...rest] = row;

            const currentChapter = chapter || '';
            let currentSection = '';
            if (subsection && subsection.trim() !== '') {
                currentSection = subsection.trim();
            } else if (section && section.trim() !== '') {
                currentSection = section.trim();
            }

            // Tạo 要件名称
            let requirementName = '';
            if (currentChapter && currentSection) {
                requirementName = `${currentChapter}_${currentSection}\n${no}`;
            } else if (currentChapter) {
                requirementName = `${currentChapter}\n${no}`;
            } else if (currentSection) {
                requirementName = `${currentSection}\n${no}`;
            } else {
                requirementName = no;
            }

            const specFileName = `要求仕様書_Weather_国内SP_Functional_Spec\n${no}`;

            // Tạo base row
            const newRow = [
                rowIndex,
                'Weather',
                currentChapter,
                currentSection,
                '',
                '',
                requirementName,
                '',
                '',
                'V7100054',
                specFileName,
                spec || '',
                tag || '',
                link || '',
                '',
                '',
                'Fix済み',
                '',
                '',
                '',
                ''
            ];

            // Mở rộng với model support
            const extendedRow = this.extendRowWithModelSupport(newRow, row);
            converted.push(extendedRow);
            rowIndex++;
        }

        return converted;
    }

    extendRowWithModelSupport(baseRow, sourceRow) {
        const extendedRow = [...baseRow];
        
        while (extendedRow.length < 100) {
            extendedRow.push('');
        }

        // Map model support từ source row
        const modelColumns = this.getModelColumns();
        for (const modelCol of modelColumns) {
            const sourceValue = sourceRow[modelCol.sourceIndex];
            const templateIndex = modelCol.templateIndex;
            
            if (templateIndex !== -1 && sourceValue) {
                let convertedValue = '';
                if (sourceValue === '○' || sourceValue === 'O' || sourceValue === 'o' || sourceValue === '0') {
                    convertedValue = '〇';
                } else if (sourceValue === '×' || sourceValue === 'X' || sourceValue === 'x') {
                    convertedValue = '×';
                } else {
                    convertedValue = sourceValue;
                }
                
                extendedRow[templateIndex] = convertedValue;
            }
        }
        
        return extendedRow;
    }

    getModelColumns() {
        const modelColumns = [];
        const headers = this.weatherData.headers;
        
        for (let i = 7; i < headers.length; i++) {
            const header = headers[i];
            if (header && header.match && header.match(/[EG]B\d{4}[VU]/)) {
                const templateIndex = this.templateData.modelColumnMap[header] || -1;
                modelColumns.push({
                    name: header,
                    sourceIndex: i,
                    templateIndex: templateIndex
                });
            }
        }
        
        return modelColumns;
    }

    performExport() {
        const newWorkbook = XLSX.utils.book_new();
        
        // Copy tất cả sheets từ template
        const originalWorkbook = this.templateData.workbook;
        originalWorkbook.SheetNames.forEach(sheetName => {
            const originalSheet = originalWorkbook.Sheets[sheetName];
            if (sheetName === this.templateData.sheetName) {
                // Thay thế sheet 要件情報 với dữ liệu mới
                const templateData = XLSX.utils.sheet_to_json(originalSheet, { header: 1 });
                const newData = [
                    ...templateData.slice(0, 6), // Giữ nguyên header và dòng mẫu
                    ...this.convertedData        // Thêm dữ liệu converted
                ];
                const newSheet = XLSX.utils.aoa_to_sheet(newData);
                XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
            } else {
                // Copy sheet khác như cũ
                const copiedSheet = XLSX.utils.aoa_to_sheet(XLSX.utils.sheet_to_json(originalSheet, { header: 1 }));
                XLSX.utils.book_append_sheet(newWorkbook, copiedSheet, sheetName);
            }
        });

        // Export file
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
        const filename = `automated_requirements_${timestamp}.xlsx`;
        XLSX.writeFile(newWorkbook, filename);
    }

    setProgress(percentage) {
        document.getElementById('autoProgressFill').style.width = percentage + '%';
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    showAutoStatus(type, message, showSpinner = false) {
        const status = document.getElementById('autoStatus');
        const statusText = document.getElementById('autoStatusText');
        const spinner = document.getElementById('autoSpinner');
        
        status.className = `status ${type}`;
        status.style.display = 'flex';
        statusText.textContent = message;
        spinner.style.display = showSpinner ? 'block' : 'none';
        
        if (!showSpinner) {
            setTimeout(() => {
                status.style.display = 'none';
            }, 5000);
        }
    }

    clearAutomation() {
        this.weatherData = null;
        this.templateData = null;
        this.flattenedData = null;
        this.convertedData = null;
        
        document.getElementById('autoWeatherInput').value = '';
        document.getElementById('autoTemplateInput').value = '';
        document.getElementById('autoWeatherUpload').classList.remove('loaded');
        document.getElementById('autoTemplateUpload').classList.remove('loaded');
        document.getElementById('autoWeatherInfo').style.display = 'none';
        document.getElementById('autoTemplateInfo').style.display = 'none';
        document.getElementById('autoProcessBtn').disabled = true;
        document.getElementById('autoStatus').style.display = 'none';
        document.getElementById('autoProgress').style.display = 'none';
        this.setProgress(0);
        
        this.showAutoStatus('info', 'Đã xóa tất cả dữ liệu');
    }
}

// Global functions
function openTool(toolType) {
    if (toolType === 'flatten') {
        window.open('excel-tree-flattener.html', '_blank');
    } else if (toolType === 'converter') {
        window.open('spec-to-requirements-converter.html', '_blank');
    }
}

function runAutomation() {
    window.automationSuite.runAutomation();
}

function clearAutomation() {
    window.automationSuite.clearAutomation();
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.automationSuite = new AutomationSuite();
    
    // Add hover effects to tool cards
    document.querySelectorAll('.tool-card').forEach(card => {
        card.addEventListener('click', () => {
            const toolId = card.id;
            if (toolId === 'tool1') {
                openTool('flatten');
            } else if (toolId === 'tool2') {
                openTool('converter');
            }
        });
    });
});