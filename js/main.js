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

        this.showAutoStatus('info', 'Đang đọc Functional Spec file...', true);

        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, {
                cellStyles: true,
                cellFormulas: true,
                cellDates: true,
                cellNF: true,
                sheetStubs: true,
                type: 'array'
            });

            console.log('Workbook sheets:', workbook.SheetNames);

            // Tìm sheet Functional_Spec - cải thiện logic tìm kiếm
            let functionalSpecSheet = null;

            // Tìm exact match trước
            for (const name of workbook.SheetNames) {
                if (name.toLowerCase().includes('functional_spec_flattened') ||
                    name.toLowerCase().includes('functional') && name.toLowerCase().includes('flattened')) {
                    functionalSpecSheet = name;
                    break;
                }
            }

            // Nếu không tìm thấy, tìm functional spec thông thường
            if (!functionalSpecSheet) {
                functionalSpecSheet = workbook.SheetNames.find(name =>
                    name.toLowerCase().includes('functional') ||
                    name.toLowerCase().includes('spec') ||
                    name.includes('仕様')
                );
            }

            // Nếu vẫn không tìm thấy, lấy sheet đầu tiên
            if (!functionalSpecSheet && workbook.SheetNames.length > 0) {
                functionalSpecSheet = workbook.SheetNames[0];
                console.warn('Không tìm thấy sheet phù hợp, sử dụng sheet đầu tiên:', functionalSpecSheet);
            }

            if (!functionalSpecSheet) {
                throw new Error('Không tìm thấy sheet nào trong file Excel');
            }

            console.log('Using sheet:', functionalSpecSheet);

            const worksheet = workbook.Sheets[functionalSpecSheet];

            // Kiểm tra worksheet có tồn tại không
            if (!worksheet) {
                throw new Error(`Sheet "${functionalSpecSheet}" không tồn tại`);
            }

            // Kiểm tra range
            if (!worksheet['!ref']) {
                throw new Error('Sheet không có dữ liệu (missing !ref)');
            }

            const range = XLSX.utils.decode_range(worksheet['!ref']);
            console.log('Sheet range:', range);

            // Đọc tất cả dữ liệu với error handling
            let allData = [];
            for (let R = range.s.r; R <= range.e.r; ++R) {
                let rowData = [];
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({c: C, r: R});
                    const cell = worksheet[cellAddress];
                    // Đảm bảo luôn có giá trị, tránh undefined
                    const cellValue = cell ? (cell.v !== undefined ? cell.v : '') : '';
                    rowData.push(cellValue);
                }
                allData.push(rowData);
            }

            console.log('Total rows read:', allData.length);
            console.log('First 3 rows:', allData.slice(0, 3));

            // Tìm header row - cải thiện logic tìm kiếm
            let headerRowIndex = -1;
            for (let i = 0; i < Math.min(10, allData.length); i++) {
                const row = allData[i];
                if (row && row.length > 0) {
                    // Kiểm tra nhiều pattern để tìm header
                    if ((row[0] && row[0].toString().toLowerCase().includes('no')) &&
                        (row[1] && row[1].toString().toLowerCase().includes('chapter'))) {
                        headerRowIndex = i;
                        break;
                    }
                }
            }

            if (headerRowIndex === -1) {
                // Nếu không tìm thấy header theo pattern, thử dùng dòng đầu tiên có dữ liệu
                for (let i = 0; i < Math.min(5, allData.length); i++) {
                    if (allData[i] && allData[i].length > 5) {
                        headerRowIndex = i;
                        console.warn('Không tìm thấy header theo pattern, sử dụng dòng:', i);
                        break;
                    }
                }
            }

            if (headerRowIndex === -1 || !allData[headerRowIndex]) {
                throw new Error('Không tìm thấy header row hợp lệ trong file');
            }

            const headers = allData[headerRowIndex] || [];
            console.log('Headers found at row', headerRowIndex, ':', headers.slice(0, 10));

            // Lấy data rows
            const dataRows = allData.slice(headerRowIndex + 1) || [];

            // Lọc các dòng có dữ liệu hợp lệ
            const weatherRows = dataRows.filter(row => {
                // Đảm bảo row tồn tại và là array
                if (!row || !Array.isArray(row)) return false;

                // Kiểm tra có ID không (column 0) và có nội dung spec không (column 4 hoặc cột có spec)
                const hasId = row[0] && row[0].toString().trim() !== '';
                const hasSpec = row[4] && row[4].toString().trim() !== '';

                return hasId || hasSpec; // Chấp nhận nếu có ID hoặc có spec
            });

            console.log(`Filtered ${weatherRows.length} valid rows from ${dataRows.length} total rows`);

            // Đảm bảo có dữ liệu
            if (weatherRows.length === 0) {
                throw new Error('Không tìm thấy dòng dữ liệu hợp lệ nào trong file');
            }

            this.weatherData = {
                headers: headers,
                rows: weatherRows
            };

            document.getElementById('autoWeatherUpload').classList.add('loaded');
            document.getElementById('autoWeatherInfo').style.display = 'block';
            document.getElementById('autoWeatherInfo').textContent =
                `✅ Đã load ${weatherRows.length} functional requirements`;

            this.showAutoStatus('success', `Đã đọc ${weatherRows.length} functional requirements`);
            this.checkAutoReady();

        } catch (error) {
            console.error('Error in handleWeatherFile:', error);
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

            console.log('Template workbook sheets:', workbook.SheetNames);

            // Tìm sheet 要件情報
            const requirementsSheet = workbook.SheetNames.find(name =>
                name.includes('要件情報') || name.includes('要件')
            );

            if (!requirementsSheet) {
                throw new Error('Không tìm thấy sheet 要件情報');
            }

            console.log('Using template sheet:', requirementsSheet);

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
            console.error('Error in handleTemplateFile:', error);
            this.showAutoStatus('error', `Lỗi đọc template file: ${error.message}`);
        }
    }

    buildTemplateModelMap(worksheet) {
        const modelMap = {};

        try {
            if (!worksheet || !worksheet['!ref']) {
                console.warn('Template worksheet không hợp lệ');
                return modelMap;
            }

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

            console.log('Built template model map:', modelMap);
        } catch (error) {
            console.error('Error building template model map:', error);
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
        if (!this.weatherData || !this.templateData) {
            this.showAutoStatus('error', 'Vui lòng upload đủ cả 2 file trước khi chạy automation');
            return;
        }

        try {
            this.setProgress(0);
            document.getElementById('autoProgress').style.display = 'block';

            // Step 1: Flatten
            this.showAutoStatus('info', 'Bước 1/3: Đang flatten cấu trúc tree...', true);
            this.setProgress(20);
            await this.sleep(500);

            this.flattenedData = this.performFlatten();
            console.log('Flattened data count:', this.flattenedData ? this.flattenedData.length : 0);
            this.setProgress(40);

            // Step 2: Convert
            this.showAutoStatus('info', 'Bước 2/3: Đang convert sang requirements format...', true);
            await this.sleep(500);

            this.convertedData = this.performConvert();
            console.log('Converted data count:', this.convertedData ? this.convertedData.length : 0);
            this.setProgress(80);

            // Step 3: Export
            this.showAutoStatus('info', 'Bước 3/3: Đang export file Excel...', true);
            await this.sleep(500);

            this.performExport();
            this.setProgress(100);

            this.showAutoStatus('success', '🎉 Hoàn thành! File đã được tải về thành công.');

        } catch (error) {
            console.error('Error in runAutomation:', error);
            this.showAutoStatus('error', `Lỗi trong quá trình xử lý: ${error.message}`);
        }
    }

    performFlatten() {
        try {
            if (!this.weatherData || !this.weatherData.rows) {
                throw new Error('Không có dữ liệu weather để flatten');
            }

            const flattened = [];
            let currentChapter = '';
            let currentSection = '';
            let currentSubsection = '';

            // Lưu trữ link và tag của từng level
            let chapterLink = '', chapterTag = '';
            let sectionLink = '', sectionTag = '';
            let subsectionLink = '', subsectionTag = '';

            for (const row of this.weatherData.rows) {
                // Đảm bảo row là array và có đủ phần tử
                if (!row || !Array.isArray(row)) continue;

                // Destructuring an toàn với default values
                const no = row[0] || '';
                const chapter = row[1] || '';
                const section = row[2] || '';
                const subsection = row[3] || '';
                const spec = row[4] || '';
                const link = row[5] || '';
                const tag = row[6] || '';
                const rest = row.slice(7) || [];

                // Cập nhật Chapter level
                if (chapter && chapter.toString().trim() !== '') {
                    currentChapter = chapter.toString().trim();
                    chapterLink = link.toString() || '';
                    chapterTag = tag.toString() || '';
                    // Reset lower levels
                    currentSection = '';
                    currentSubsection = '';
                    sectionLink = '';
                    sectionTag = '';
                    subsectionLink = '';
                    subsectionTag = '';
                }

                // Cập nhật Section level
                if (section && section.toString().trim() !== '') {
                    currentSection = section.toString().trim();
                    sectionLink = link.toString() || '';
                    sectionTag = tag.toString() || '';
                    // Reset lower level
                    currentSubsection = '';
                    subsectionLink = '';
                    subsectionTag = '';
                }

                // Cập nhật Subsection level
                if (subsection && subsection.toString().trim() !== '') {
                    currentSubsection = subsection.toString().trim();
                    subsectionLink = link.toString() || '';
                    subsectionTag = tag.toString() || '';
                }

                // Chỉ xử lý dòng có functional specification
                if (spec && spec.toString().trim() !== '') {
                    // Gộp link từ tất cả levels - filter empty strings an toàn
                    const allLinks = [
                        chapterLink,
                        sectionLink,
                        subsectionLink,
                        link.toString() || ''
                    ].filter(l => l && l.trim() !== '').join('\n');

                    // Gộp tag từ tất cả levels - filter empty strings an toàn
                    const allTags = [
                        chapterTag,
                        sectionTag,
                        subsectionTag,
                        tag.toString() || ''
                    ].filter(t => t && t.trim() !== '').join('\n');

                    const flatRow = [
                        no.toString(),
                        currentChapter,
                        currentSection,
                        currentSubsection,
                        spec.toString(),
                        allLinks,    // Gộp tất cả links
                        allTags,     // Gộp tất cả tags
                        ...rest      // Spread an toàn
                    ];

                    flattened.push(flatRow);
                }
            }

            console.log('Flatten completed:', flattened.length, 'rows processed');
            return flattened;

        } catch (error) {
            console.error('Error in performFlatten:', error);
            throw new Error(`Lỗi flatten dữ liệu: ${error.message}`);
        }
    }

    performConvert() {
        try {
            if (!this.flattenedData || !Array.isArray(this.flattenedData)) {
                throw new Error('Không có dữ liệu flattened để convert');
            }

            const converted = [];
            let rowIndex = 1;

            // Lọc chỉ các dòng có Functional Specification - với error handling
            const specRows = this.flattenedData.filter(row => {
                if (!row || !Array.isArray(row)) return false;
                const spec = row[4];
                return spec && spec.toString().trim() !== '';
            });

            console.log(`Converting ${specRows.length} specification rows`);

            for (const row of specRows) {
                // Destructuring an toàn
                const no = row[0] || '';
                const chapter = row[1] || '';
                const section = row[2] || '';
                const subsection = row[3] || '';
                const spec = row[4] || '';
                const link = row[5] || '';
                const tag = row[6] || '';
                const rest = row.slice(7) || [];

                // Vì file đã flatten nên chapter và section đã có sẵn ở mọi dòng
                const currentChapter = chapter.toString() || '';

                // Ưu tiên subsection, nếu không có thì lấy section
                let currentSection = '';
                if (subsection && subsection.toString().trim() !== '') {
                    currentSection = subsection.toString().trim();
                } else if (section && section.toString().trim() !== '') {
                    currentSection = section.toString().trim();
                }

                // Tạo 要件ID từ cột No.
                const requirementId = no.toString() || '';

                // Tạo 要件名称 theo format mới: chapter_section_subsection (không có No.)
                let requirementName = '';
                const nameParts = [];
                if (currentChapter && currentChapter.trim() !== '') {
                    nameParts.push(currentChapter.trim());
                }
                if (section && section.toString().trim() !== '') {
                    nameParts.push(section.toString().trim());
                }
                if (subsection && subsection.toString().trim() !== '') {
                    nameParts.push(subsection.toString().trim());
                }
                requirementName = nameParts.join('_');

                // Tạo 仕様書ファイル名 theo format mới - EXTRACT APP NAME từ ID
                let appName = 'App';

                // Thử extract app name từ ID pattern
                const noStr = no.toString();
                if (noStr.includes('_')) {
                    const prefix = noStr.split('_')[0];
                    if (prefix && prefix.length >= 3) {
                        // Map common prefixes
                        const appMap = {
                            'WEA': 'Weather',
                            'PED': 'Pedometer',
                            'CAL': 'Calendar',
                            'CAM': 'Camera',
                            'GAL': 'Gallery',
                            'MUS': 'Music',
                            'VID': 'Video',
                            'MSG': 'Message',
                            'PHO': 'Phone',
                            'CON': 'Contact',
                            'FLA': 'Flashlight'  // Thêm cho file hiện tại
                        };
                        appName = appMap[prefix] || prefix;
                    }
                }

                const specFileName = `要求仕様書_${appName}_国内SP_Functional_Spec_no.${noStr}`;

                // Tạo base row với error handling
                const newRow = [
                    rowIndex,                      // 0: No
                    appName,                       // 1: 機能名称 (auto-detect từ ID)
                    currentChapter,                // 2: 章
                    currentSection,                // 3: 節
                    requirementId,                 // 4: 要件ID (NEW)
                    '',                           // 5: SubID
                    requirementName,               // 6: 要件名称 (UPDATED)
                    '',                           // 7: 要件種別
                    '',                           // 8: 要求元
                    '',                           // 9: 仕様書バージョン
                    specFileName,                 // 10: 仕様書ファイル名 (UPDATED)
                    spec.toString() || '',        // 11: 要件内容
                    tag.toString() || '',         // 12: ラベル
                    link.toString() || '',        // 13: 備考
                    '',                           // 14: 要件添付ファイルパス
                    '',                           // 15: 要件原文ファイル添付ファイルパス
                    'Fix済み',                      // 16: 要件ステータス
                    '',                           // 17: コミット管理情報ID
                    '',                           // 18: コミット管理情報名称
                    '',                           // 19: コミット管理情報内容
                    ''                            // 20: EB1190V_E
                ];

                // Mở rộng với model support
                const extendedRow = this.extendRowWithModelSupport(newRow, row);
                converted.push(extendedRow);
                rowIndex++;
            }

            console.log('Convert completed:', converted.length, 'rows created');
            return converted;

        } catch (error) {
            console.error('Error in performConvert:', error);
            throw new Error(`Lỗi convert dữ liệu: ${error.message}`);
        }
    }

    extendRowWithModelSupport(baseRow, sourceRow) {
        try {
            // Đảm bảo baseRow và sourceRow là arrays
            if (!Array.isArray(baseRow)) baseRow = [];
            if (!Array.isArray(sourceRow)) sourceRow = [];

            const extendedRow = [...baseRow];

            // Mở rộng row đến 100 cột
            while (extendedRow.length < 100) {
                extendedRow.push('');
            }

            // Map model support từ source row
            const modelColumns = this.getModelColumns();

            if (modelColumns && Array.isArray(modelColumns)) {
                for (const modelCol of modelColumns) {
                    try {
                        const sourceValue = sourceRow[modelCol.sourceIndex];
                        const templateIndex = modelCol.templateIndex;

                        if (templateIndex !== -1 && sourceValue !== undefined && sourceValue !== null) {
                            let convertedValue = '';
                            const sourceStr = sourceValue.toString();

                            if (sourceStr === '○' || sourceStr === 'O' || sourceStr === 'o' || sourceStr === '0') {
                                convertedValue = '〇';
                            } else if (sourceStr === '×' || sourceStr === 'X' || sourceStr === 'x') {
                                convertedValue = '×';
                            } else {
                                convertedValue = sourceStr;
                            }

                            if (templateIndex < extendedRow.length) {
                                extendedRow[templateIndex] = convertedValue;
                            }
                        }
                    } catch (modelError) {
                        console.warn('Error processing model column:', modelCol, modelError);
                    }
                }
            }

            return extendedRow;

        } catch (error) {
            console.error('Error in extendRowWithModelSupport:', error);
            return baseRow || [];
        }
    }

    getModelColumns() {
        try {
            const modelColumns = [];

            if (!this.weatherData || !this.weatherData.headers || !Array.isArray(this.weatherData.headers)) {
                console.warn('Weather data headers not available');
                return modelColumns;
            }

            const headers = this.weatherData.headers;

            for (let i = 7; i < headers.length; i++) {
                const header = headers[i];
                if (header && typeof header === 'string') {
                    const modelMatch = header.match(/[EG]B\d{4}[VU]/);
                    if (modelMatch) {
                        let templateIndex = -1;

                        if (this.templateData && this.templateData.modelColumnMap) {
                            templateIndex = this.templateData.modelColumnMap[header] || -1;
                        }

                        modelColumns.push({
                            name: header,
                            sourceIndex: i,
                            templateIndex: templateIndex
                        });
                    }
                }
            }

            console.log('Found model columns:', modelColumns.length);
            return modelColumns;

        } catch (error) {
            console.error('Error in getModelColumns:', error);
            return [];
        }
    }

    performExport() {
        try {
            if (!this.templateData || !this.templateData.workbook) {
                throw new Error('Template data không hợp lệ');
            }

            if (!this.convertedData || !Array.isArray(this.convertedData) || this.convertedData.length === 0) {
                throw new Error('Không có dữ liệu converted để export');
            }

            const newWorkbook = XLSX.utils.book_new();

            // Copy tất cả sheets từ template
            const originalWorkbook = this.templateData.workbook;

            if (!originalWorkbook.SheetNames || !Array.isArray(originalWorkbook.SheetNames)) {
                throw new Error('Template workbook không có sheets');
            }

            originalWorkbook.SheetNames.forEach(sheetName => {
                try {
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
                        const copiedData = XLSX.utils.sheet_to_json(originalSheet, { header: 1 });
                        const copiedSheet = XLSX.utils.aoa_to_sheet(copiedData);
                        XLSX.utils.book_append_sheet(newWorkbook, copiedSheet, sheetName);
                    }
                } catch (sheetError) {
                    console.error(`Error processing sheet ${sheetName}:`, sheetError);
                    // Skip problematic sheets
                }
            });

            // Export file
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
            const filename = `automated_requirements_v2_${timestamp}.xlsx`;
            XLSX.writeFile(newWorkbook, filename);

            console.log('Export completed:', filename);

        } catch (error) {
            console.error('Error in performExport:', error);
            throw new Error(`Lỗi export file: ${error.message}`);
        }
    }

    setProgress(percentage) {
        try {
            const progressFill = document.getElementById('autoProgressFill');
            if (progressFill) {
                progressFill.style.width = percentage + '%';
            }
        } catch (error) {
            console.error('Error setting progress:', error);
        }
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    showAutoStatus(type, message, showSpinner = false) {
        try {
            const status = document.getElementById('autoStatus');
            const statusText = document.getElementById('autoStatusText');
            const spinner = document.getElementById('autoSpinner');

            if (status && statusText) {
                status.className = `status ${type}`;
                status.style.display = 'flex';
                statusText.textContent = message;

                if (spinner) {
                    spinner.style.display = showSpinner ? 'block' : 'none';
                }

                if (!showSpinner) {
                    setTimeout(() => {
                        status.style.display = 'none';
                    }, 5000);
                }
            }
        } catch (error) {
            console.error('Error showing status:', error);
        }
    }

    clearAutomation() {
        this.weatherData = null;
        this.templateData = null;
        this.flattenedData = null;
        this.convertedData = null;

        try {
            const elements = [
                'autoWeatherInput',
                'autoTemplateInput',
                'autoWeatherUpload',
                'autoTemplateUpload',
                'autoWeatherInfo',
                'autoTemplateInfo',
                'autoProcessBtn',
                'autoStatus',
                'autoProgress'
            ];

            elements.forEach(id => {
                const element = document.getElementById(id);
                if (element) {
                    switch(id) {
                        case 'autoWeatherInput':
                        case 'autoTemplateInput':
                            element.value = '';
                            break;
                        case 'autoWeatherUpload':
                        case 'autoTemplateUpload':
                            element.classList.remove('loaded');
                            break;
                        case 'autoWeatherInfo':
                        case 'autoTemplateInfo':
                        case 'autoStatus':
                        case 'autoProgress':
                            element.style.display = 'none';
                            break;
                        case 'autoProcessBtn':
                            element.disabled = true;
                            break;
                    }
                }
            });

            this.setProgress(0);
            this.showAutoStatus('info', 'Đã xóa tất cả dữ liệu');

        } catch (error) {
            console.error('Error clearing automation:', error);
        }
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
    if (window.automationSuite) {
        window.automationSuite.runAutomation();
    } else {
        console.error('AutomationSuite not initialized');
    }
}

function clearAutomation() {
    if (window.automationSuite) {
        window.automationSuite.clearAutomation();
    } else {
        console.error('AutomationSuite not initialized');
    }
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    try {
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

        console.log('AutomationSuite initialized successfully');
    } catch (error) {
        console.error('Error initializing AutomationSuite:', error);
    }
});