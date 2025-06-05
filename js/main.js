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

            // Tìm header row - tìm trong nhiều dòng đầu
            let headerRowIndex = -1;
            for (let i = 0; i < Math.min(10, allData.length); i++) {
                if (allData[i][0] === 'No.' && allData[i][1] === 'Chapter') {
                    headerRowIndex = i;
                    break;
                }
            }

            const headers = allData[headerRowIndex];
            const dataRows = allData.slice(headerRowIndex + 1);
            const weatherRows = dataRows.filter(row => {
                // Chỉ cần check có ID không (không rỗng và có nội dung)
                return row[0] && row[0].toString().trim() !== '';
            });

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

      // Lưu trữ link và tag của từng level
      let chapterLink = '', chapterTag = '';
      let sectionLink = '', sectionTag = '';
      let subsectionLink = '', subsectionTag = '';

      for (const row of this.weatherData.rows) {
          const [no, chapter, section, subsection, spec, link, tag, ...rest] = row;

          // Cập nhật Chapter level
          if (chapter && chapter.trim() !== '') {
              currentChapter = chapter.trim();
              chapterLink = link || '';
              chapterTag = tag || '';
              // Reset lower levels
              currentSection = '';
              currentSubsection = '';
              sectionLink = '';
              sectionTag = '';
              subsectionLink = '';
              subsectionTag = '';
          }

          // Cập nhật Section level
          if (section && section.trim() !== '') {
              currentSection = section.trim();
              sectionLink = link || '';
              sectionTag = tag || '';
              // Reset lower level
              currentSubsection = '';
              subsectionLink = '';
              subsectionTag = '';
          }

          // Cập nhật Subsection level
          if (subsection && subsection.trim() !== '') {
              currentSubsection = subsection.trim();
              subsectionLink = link || '';
              subsectionTag = tag || '';
          }

          // Chỉ xử lý dòng có functional specification
          if (spec && spec.trim() !== '') {
              // Gộp link từ tất cả levels
              const allLinks = [
                  chapterLink,
                  sectionLink,
                  subsectionLink,
                  link || ''
              ].filter(l => l && l.trim() !== '').join('\n');

              // Gộp tag từ tất cả levels
              const allTags = [
                  chapterTag,
                  sectionTag,
                  subsectionTag,
                  tag || ''
              ].filter(t => t && t.trim() !== '').join('\n');

              const flatRow = [
                  no,
                  currentChapter,
                  currentSection,
                  currentSubsection,
                  spec,
                  allLinks,    // Gộp tất cả links
                  allTags,     // Gộp tất cả tags
                  ...rest
              ];

              flattened.push(flatRow);
          }
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

                                // Vì file đã flatten nên chapter và section đã có sẵn ở mọi dòng
                                const currentChapter = chapter || '';

                                // Ưu tiên subsection, nếu không có thì lấy section
                                let currentSection = '';
                                if (subsection && subsection.trim() !== '') {
                                    currentSection = subsection.trim();
                                } else if (section && section.trim() !== '') {
                                    currentSection = section.trim();
                                }

                                // Tạo 要件ID từ cột No. (WEA_1.1.1.1)
                                const requirementId = no || '';

                                // Tạo 要件名称 theo format mới: chapter_section_subsection (không có No.)
                                let requirementName = '';
                                const nameParts = [];
                                if (currentChapter && currentChapter.trim() !== '') {
                                    nameParts.push(currentChapter.trim());
                                }
                                if (section && section.trim() !== '') {
                                    nameParts.push(section.trim());
                                }
                                if (subsection && subsection.trim() !== '') {
                                    nameParts.push(subsection.trim());
                                }
                                requirementName = nameParts.join('_');

            // Tạo 仕様書ファイル名 theo format mới - EXTRACT APP NAME từ ID
            let appName = 'App';

            // Thử extract app name từ ID pattern
            if (no.includes('_')) {
                const prefix = no.split('_')[0];
                if (prefix.length >= 3) {
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
                        'CON': 'Contact'
                    };
                    appName = appMap[prefix] || prefix;
                }
            }

            const specFileName = `要求仕様書_${appName}_国内SP_Functional_Spec_no.${no}`;

            // Tạo base row
            const newRow = [
                rowIndex,                  // 0: No
                appName,                   // 1: 機能名称 (auto-detect từ ID)
                currentChapter,            // 2: 章
                currentSection,            // 3: 節
                requirementId,             // 4: 要件ID (NEW)
                '',                        // 5: SubID
                requirementName,           // 6: 要件名称 (UPDATED)
                '',                        // 7: 要件種別
                '',                        // 8: 要求元
                '',               // 9: 仕様書バージョン
                specFileName,             // 10: 仕様書ファイル名 (UPDATED)
                spec || '',               // 11: 要件内容
                tag || '',                // 12: ラベル
                link || '',               // 13: 備考
                '',                        // 14: 要件添付ファイルパス
                '',                        // 15: 要件原文ファイル添付ファイルパス
                'Fix済み',                   // 16: 要件ステータス
                '',                        // 17: コミット管理情報ID
                '',                        // 18: コミット管理情報名称
                '',                        // 19: コミット管理情報内容
                ''                         // 20: EB1190V_E
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
        const filename = `automated_requirements_v2_${timestamp}.xlsx`;
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