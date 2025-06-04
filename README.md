# Android Development Excel Tools

## Mô tả
Bộ công cụ Excel chuyên dụng cho phát triển Android, bao gồm:
1. **Excel Tree Flattener** - Chuyển đổi cấu trúc tree thành flat
2. **Spec to Requirements Converter** - Chuyển đổi từ Functional Spec sang 要件情報 format

## Tool 1: Excel Tree Flattener

### Tính năng
- ✅ Đọc file Excel (.xlsx, .xls) 
- ✅ Tự động tìm sheet "Functional_Spec"
- ✅ Flatten cấu trúc tree thành flat structure
- ✅ Preview dữ liệu trước và sau khi convert
- ✅ Export file Excel kết quả
- ✅ Giao diện thân thiện với drag & drop

## Tool 2: Spec to Requirements Converter

### Tính năng
- ✅ Chuyển đổi từ Weather Functional Spec sang 要件情報 format
- ✅ Tự động mapping theo quy tắc định sẵn
- ✅ Chỉ lấy items có "機能仕様 / Functional Specification"
- ✅ Tạo 要件名称 theo format: "Chapter_Section\nWEA_No."
- ✅ Tạo 仕様書ファイル名 theo format chuẩn
- ✅ Preview kết quả trước khi export
- ✅ Giữ nguyên template structure

## Cách sử dụng

### Tool 1: Excel Tree Flattener
```bash
# Mở file trong trình duyệt
D:\dev\tools\excel-tree-flattener.html
```

### Tool 2: Spec to Requirements Converter
```bash
# Mở file trong trình duyệt
D:\dev\tools\spec-to-requirements-converter.html
```

### Sử dụng Tool 1 (Tree Flattener):
1. Upload file Weather Functional Spec
2. Click "Flatten Tree Structure"
3. Preview kết quả
4. Export file Excel mới

### Sử dụng Tool 2 (Requirements Converter):
1. Upload file Weather Functional Spec (source)
2. Upload file Requirements Template (target)
3. Click "Convert to Requirements"
4. Preview kết quả theo format 要件情報
5. Export file Excel mới

## Cấu trúc dữ liệu

### Tool 1: Tree Flattener

**Input (Tree Structure):**
```
WEA_1.0.0.0 | Weather application |              |                    | 
WEA_1.1.0.0 |                    | 天気データ   |                    |
WEA_1.1.1.0 |                    |              | 提供元             |
WEA_1.1.1.1 |                    |              |                    | AccuWeatherから天気データを取得する
```

**Output (Flat Structure):**
```
WEA_1.0.0.0 | Weather application |              |                    |
WEA_1.1.0.0 | Weather application | 天気データ   |                    |
WEA_1.1.1.0 | Weather application | 天気データ   | 提供元             |
WEA_1.1.1.1 | Weather application | 天気データ   | 提供元             | AccuWeatherから天気データを取得する
```

### Tool 2: Requirements Converter

**Mapping Rules:**
- Section → 節
- Chapter → 章  
- 要件名称 = "Chapter" + "_" + "Section" + "\n" + "WEA_No."
- Tag → ラベル
- Link → 備考
- 仕様書ファイル名 = "要求仕様書_Weather_国内SP_Functional_Spec\nWEA_[No.]"

**Example Output:**
```
No. | 機能名称 | 章 | 節 | 要件名称
1   | Weather | Weather application | 天気データ Weather info | Weather application_天気データ Weather info WEA_1.1.1.1
```

## Quy tắc xử lý

### Tool 1 (Tree Flattener):
1. **Chapter**: Khi gặp dòng có Chapter mới, reset Section và Subsection
2. **Section**: Khi gặp dòng có Section mới, reset Subsection  
3. **Subsection**: Giữ nguyên cho đến khi có Section hoặc Chapter mới
4. **Specification**: Luôn lấy từ dòng hiện tại

### Tool 2 (Requirements Converter):
1. **Filter**: Chỉ lấy rows có "機能仕様 / Functional Specification" không rỗng
2. **Hierarchy**: Xây dựng lại hierarchy từ tree structure
3. **Mapping**: Áp dụng mapping rules để tạo format 要件情報
4. **Template**: Giữ nguyên structure của template file

## Yêu cầu hệ thống

- Trình duyệt web hiện đại (Chrome, Firefox, Edge, Safari)
- Hỗ trợ JavaScript ES6+
- Không cần cài đặt thêm phần mềm

## Thư viện sử dụng

- **SheetJS (XLSX)**: Đọc/ghi file Excel
- **Vanilla JavaScript**: Logic xử lý
- **CSS3**: Giao diện responsive

## Files trong thư mục

- `excel-tree-flattener.html` - Tool chính cho flatten tree
- `spec-to-requirements-converter.html` - Tool convert sang requirements format
- `run-tool.bat` - Shortcut để mở tools
- `README.md` - Tài liệu hướng dẫn

## Tác giả
Tools được phát triển cho việc xử lý tài liệu functional specification trong phát triển ứng dụng Android.

## Phiên bản
- v1.0.0 - Excel Tree Flattener
- v1.1.0 - Thêm Spec to Requirements Converter
- Hỗ trợ format WEA_ và 要件情報

## Lưu ý
- Tool 1 chỉ xử lý các dòng có mã bắt đầu bằng "WEA_"
- Tool 2 chỉ convert items có Functional Specification
- Cần có header row chứa "No." và "Chapter"
- File kết quả sẽ giữ nguyên format Excel gốc