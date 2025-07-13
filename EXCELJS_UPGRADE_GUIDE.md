# 🚀 ExcelJS Upgrade Guide

## 🎯 **Tại sao nên upgrade lên ExcelJS?**

### ❌ **Vấn đề hiện tại với XLSX.js:**
- ❓ **Colors bị mất**: Chỉ hiển thị `{"patternType":"none"}`
- ❓ **Borders bị mất**: Không preserve được border styling  
- ❓ **Font formatting bị mất**: Bold, italic, font size không được giữ
- ❓ **Limited compatibility**: Conditional formatting không support

### ✅ **ExcelJS sẽ giải quyết:**
- ✅ **Full color support**: RGB values, theme colors
- ✅ **Complete border styling**: All border types và styles
- ✅ **Rich font formatting**: Bold, italic, size, family, colors
- ✅ **Conditional formatting**: Read và preserve
- ✅ **Better Excel compatibility**: Professional output

---

## 📊 **So sánh chi tiết:**

| Feature | XLSX.js | ExcelJS |
|---------|---------|---------|
| **Color Support** | ❓ Simplified | ✅ Full RGB + Themes |
| **Border Styling** | ❓ Basic | ✅ Complete |
| **Font Formatting** | ❓ Limited | ✅ Rich support |
| **Conditional Formatting** | ❌ No | ✅ Yes |
| **Performance** | ⚡ Fast | ⚡ Moderate |
| **File Size** | 📊 Smaller | 📊 Larger |
| **API Complexity** | 📚 Simple | 📚 Rich |
| **Best For** | Data processing | Rich formatting |

---

## 🔧 **Cách upgrade:**

### **Step 1: Install ExcelJS**
```bash
npm install exceljs
npm install @types/exceljs  # for TypeScript
```

### **Step 2: Update imports trong page.tsx**
```typescript
// OLD import
import * as XLSX from "xlsx";
import {
  deepCopyWorkbook,
  detectFileFormatting,
  // ... other utils
} from "@/lib/excel-utils";

// NEW import  
import { createBestExcelHandler } from "@/lib/excel-handler";
```

### **Step 3: Replace handleFileChange**
```typescript
// OLD implementation
const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  const file = event.target.files?.[0];
  if (!file) return;
  
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target?.result as ArrayBuffer);
    const wb = XLSX.read(data, { /* options */ });
    // ... processing
  };
  reader.readAsArrayBuffer(file);
};

// NEW implementation
const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
  const file = event.target.files?.[0];
  if (!file) return;
  
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      
      // Create universal handler (auto-detects ExcelJS/XLSX.js)
      const handler = await createBestExcelHandler();
      const result = await handler.readFile(data);
      
      console.log(`Using library: ${result.library}`);
      console.log(`Formatting detected: ${result.hasFormatting}`);
      
      setSheetNames(result.sheets);
      
      if (result.sheets.length === 1) {
        const sheetData = handler.processSheetData(result.sheets[0]);
        setHeaders(sheetData.headers);
        setRows(sheetData.rows);
        setActiveSheetName(result.sheets[0]);
      } else {
        setIsSheetSelectorOpen(true);
      }
      
      // Store handler for export
      setExcelHandler(handler);
      
      toast({
        title: "Success!",
        description: `File loaded with ${result.library}. Formatting: ${result.hasFormatting ? 'Detected' : 'Basic'}`,
      });
      
    } catch (error) {
      console.error("Error processing Excel file:", error);
      toast({
        variant: "destructive",
        title: "File Error", 
        description: "Could not process the Excel file.",
      });
    }
  };
  reader.readAsArrayBuffer(file);
};
```

### **Step 4: Replace handleExport**
```typescript
// OLD implementation  
const handleExport = () => {
  if (!originalFileData || !activeSheetName) return;
  
  // Complex XLSX.js logic...
  const originalWorkbook = XLSX.read(originalFileData, { /* options */ });
  // ... lots of manual copying and formatting logic
};

// NEW implementation
const handleExport = async () => {
  if (!excelHandler || !activeSheetName) {
    toast({
      variant: "destructive",
      title: "No Data Found",
      description: "Please upload an Excel file first."
    });
    return;
  }

  try {
    console.log('Starting export with enhanced formatting preservation...');
    
    // Export with automatic formatting preservation
    const exportBuffer = await excelHandler.exportWithCheckIns(
      activeSheetName,
      rows
    );
    
    // Create and download
    const blob = new Blob([exportBuffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });

    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', 'attendee_report_updated.xlsx');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    toast({
      title: "Export Successful! ✨",
      description: "The updated attendee report has been downloaded with enhanced formatting preservation.",
    });
    
  } catch (error) {
    console.error("Error exporting Excel file:", error);
    toast({
      variant: "destructive",
      title: "Export Error",
      description: "Could not export the Excel file: " + (error instanceof Error ? error.message : String(error)),
    });
  }
};
```

### **Step 5: Add state for handler**
```typescript
// Add this state
const [excelHandler, setExcelHandler] = React.useState<any>(null);
```

### **Step 6: Update processSheetData**
```typescript
// OLD implementation
const processSheetData = (wb: WorkBook, sheetName: string) => {
  const worksheet = wb.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
  // ... processing
};

// NEW implementation  
const processSheetData = (sheetName: string) => {
  if (!excelHandler) return;
  
  try {
    const sheetData = excelHandler.processSheetData(sheetName);
    setHeaders(sheetData.headers);
    setRows(sheetData.rows);
    setActiveSheetName(sheetName);
    
    toast({
      title: "Success!",
      description: `Successfully imported ${sheetData.rows.length} rows from sheet: ${sheetName}.`,
    });
  } catch (error) {
    console.error('Error processing sheet data:', error);
    toast({
      variant: "destructive",
      title: "Processing Error",
      description: "There was an error processing the sheet data.",
    });
  }
};
```

---

## 🎯 **Expected Results sau khi upgrade:**

### **Với ExcelJS (nếu available):**
- ✅ **Colors preserved**: Background colors, font colors
- ✅ **Borders intact**: All border styles và colors  
- ✅ **Fonts preserved**: Bold, italic, size, family
- ✅ **Professional output**: Exactly như file gốc
- ✅ **Rich formatting**: Conditional formatting support

### **Fallback to XLSX.js (nếu ExcelJS không có):**
- ✅ **Still functional**: App vẫn hoạt động bình thường
- ✅ **Structure preserved**: Columns, rows, merges
- ❓ **Visual simplified**: Colors/borders basic

---

## 📋 **Testing Steps:**

1. **Install ExcelJS**: `npm install exceljs`
2. **Update code** theo guide trên
3. **Test với DATA.xlsx**:
   - Upload file
   - Check console log: "Using library: ExcelJS"
   - Add check-in data  
   - Export file
   - **Open in Excel và so sánh visual formatting**

4. **Verify improvements**:
   - Colors preserved?
   - Borders intact?
   - Fonts still bold/styled?
   - Overall layout identical?

---

## 🚨 **Important Notes:**

1. **Backward compatible**: Code sẽ fallback về XLSX.js nếu ExcelJS không có
2. **Performance**: ExcelJS chậm hơn một chút nhưng output tốt hơn nhiều
3. **File size**: Output files có thể lớn hơn
4. **Dependencies**: Thêm ~2MB bundle size

---

## 🎉 **Expected Impact:**

### **Before (XLSX.js only):**
- ❓ Colors mất: `{"patternType":"none"}`
- ❓ Borders mất
- ❓ Font formatting simplified
- ✅ Structure OK

### **After (với ExcelJS):**
- ✅ **Colors perfect**: Full RGB preservation
- ✅ **Borders perfect**: All styles intact
- ✅ **Fonts perfect**: Bold, colors, sizes preserved
- ✅ **Professional result**: Như file gốc

---

## 🔧 **Troubleshooting:**

**Q: ExcelJS không install được?**
A: Check node version, try: `npm install exceljs --legacy-peer-deps`

**Q: App crash sau upgrade?**  
A: Check console errors, có thể cần update TypeScript types

**Q: Performance chậm?**
A: Normal với ExcelJS, trade-off cho better formatting

**Q: File size lớn?**
A: Expected, ExcelJS preserve more data

---

## ✅ **Summary:**

Upgrade này sẽ **dramatically improve** formatting preservation:
- **Visual formatting** từ ❓ → ✅  
- **Professional output** từ ❓ → ✅
- **User satisfaction** từ ❓ → ✅

**Worth the upgrade!** 🚀