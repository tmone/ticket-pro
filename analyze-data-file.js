// Phân tích chi tiết file DATA.xlsx để hiểu thực sự có gì
const XLSX = require('xlsx');
const fs = require('fs');

console.log('🔍 PHÂN TÍCH CHI TIẾT FILE DATA.xlsx\n');

if (!fs.existsSync('DATA.xlsx')) {
  console.log('❌ File DATA.xlsx không tồn tại!');
  process.exit(1);
}

// Đọc file với TẤT CẢ options có thể
console.log('📂 Đọc file với maximum options...');

const fileData = fs.readFileSync('DATA.xlsx');

const workbook = XLSX.read(fileData, {
  type: "buffer",
  cellStyles: true,
  cellFormula: true,
  cellDates: true,
  cellNF: true,
  sheetStubs: true,
  bookVBA: true,
  bookFiles: true,
  bookProps: true,
  bookSheets: true,
  raw: false,
  codepage: 65001
});

console.log('✅ File đã đọc\n');

// Thông tin workbook
console.log('📊 WORKBOOK INFO:');
console.log(`   Sheets: ${workbook.SheetNames.length} sheets`);
console.log(`   Sheet names: ${workbook.SheetNames.join(', ')}`);

if (workbook.Props) {
  console.log('   Properties:', workbook.Props);
}
if (workbook.Custprops) {
  console.log('   Custom properties:', workbook.Custprops);
}

// Phân tích từng sheet
workbook.SheetNames.forEach((sheetName, index) => {
  console.log(`\n🗂️  SHEET ${index + 1}: "${sheetName}"`);
  
  const sheet = workbook.Sheets[sheetName];
  
  // Basic info
  console.log(`   📐 Range: ${sheet['!ref'] || 'NO RANGE'}`);
  
  // Sheet-level formatting
  console.log('   🎨 Sheet-level formatting:');
  console.log(`      Column widths (!cols): ${sheet['!cols'] ? `YES (${sheet['!cols'].length} columns)` : 'NO'}`);
  console.log(`      Row heights (!rows): ${sheet['!rows'] ? `YES (${sheet['!rows'].length} rows)` : 'NO'}`);
  console.log(`      Merged cells (!merges): ${sheet['!merges'] ? `YES (${sheet['!merges'].length} merges)` : 'NO'}`);
  console.log(`      Auto filter (!autofilter): ${sheet['!autofilter'] ? 'YES' : 'NO'}`);
  console.log(`      Protection (!protect): ${sheet['!protect'] ? 'YES' : 'NO'}`);
  console.log(`      Margins (!margins): ${sheet['!margins'] ? 'YES' : 'NO'}`);
  
  // Print detailed column info if exists
  if (sheet['!cols']) {
    console.log('   📏 Column details:');
    sheet['!cols'].slice(0, 10).forEach((col, i) => {
      if (col) {
        console.log(`      Col ${i}: width=${col.width || 'default'}, hidden=${col.hidden || false}`);
      }
    });
    if (sheet['!cols'].length > 10) {
      console.log(`      ... and ${sheet['!cols'].length - 10} more columns`);
    }
  }
  
  // Print merge info if exists
  if (sheet['!merges']) {
    console.log('   🔗 Merge details:');
    sheet['!merges'].slice(0, 5).forEach((merge, i) => {
      const start = XLSX.utils.encode_cell(merge.s);
      const end = XLSX.utils.encode_cell(merge.e);
      console.log(`      Merge ${i + 1}: ${start}:${end}`);
    });
    if (sheet['!merges'].length > 5) {
      console.log(`      ... and ${sheet['!merges'].length - 5} more merges`);
    }
  }
  
  // Cell-level analysis
  const cellsWithStyles = [];
  const cellsWithFormulas = [];
  const cellsWithComments = [];
  const cellsWithHyperlinks = [];
  let totalCells = 0;
  
  Object.keys(sheet).forEach(key => {
    if (!key.startsWith('!')) {
      totalCells++;
      const cell = sheet[key];
      
      if (cell.s) cellsWithStyles.push(key);
      if (cell.f) cellsWithFormulas.push(key);
      if (cell.c) cellsWithComments.push(key);
      if (cell.l) cellsWithHyperlinks.push(key);
    }
  });
  
  console.log('   📊 Cell analysis:');
  console.log(`      Total cells: ${totalCells}`);
  console.log(`      Cells with styles: ${cellsWithStyles.length}`);
  console.log(`      Cells with formulas: ${cellsWithFormulas.length}`);
  console.log(`      Cells with comments: ${cellsWithComments.length}`);
  console.log(`      Cells with hyperlinks: ${cellsWithHyperlinks.length}`);
  
  // Show some sample styled cells
  if (cellsWithStyles.length > 0) {
    console.log('   🎨 Sample styled cells:');
    cellsWithStyles.slice(0, 5).forEach(cellAddr => {
      const cell = sheet[cellAddr];
      console.log(`      ${cellAddr}: value="${cell.v}" style=${JSON.stringify(cell.s)}`);
    });
    if (cellsWithStyles.length > 5) {
      console.log(`      ... and ${cellsWithStyles.length - 5} more styled cells`);
    }
  }
  
  // Sample data
  if (sheet['!ref']) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    console.log('   📋 Sample data (first 3x3):');
    
    for (let r = range.s.r; r <= Math.min(range.s.r + 2, range.e.r); r++) {
      const row = [];
      for (let c = range.s.c; c <= Math.min(range.s.c + 2, range.e.c); c++) {
        const cellAddr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[cellAddr];
        row.push(cell ? `"${cell.v}"` : 'empty');
      }
      console.log(`      Row ${r + 1}: ${row.join(', ')}`);
    }
  }
});

// Test detection functions
console.log('\n🔍 DETECTION TEST:');

const detectFileFormatting = (workbook) => {
  try {
    if (!workbook || !workbook.SheetNames) return false;

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet || typeof sheet !== 'object') continue;
      
      // Check sheet-level formatting
      if (sheet['!cols'] || sheet['!rows'] || sheet['!merges'] || 
          sheet['!protect'] || sheet['!autofilter'] || sheet['!margins']) {
        console.log(`   Found sheet-level formatting in "${sheetName}"`);
        return true;
      }
      
      // Check cell-level formatting
      for (const cellAddress of Object.keys(sheet)) {
        if (!cellAddress.startsWith('!')) {
          const cell = sheet[cellAddress];
          if (cell && typeof cell === 'object' && 
              cell !== null && !Array.isArray(cell) &&
              (cell.hasOwnProperty('s') || cell.hasOwnProperty('z') || 
               cell.hasOwnProperty('l') || cell.hasOwnProperty('c'))) {
            console.log(`   Found cell-level formatting in "${sheetName}" at ${cellAddress}`);
            return true;
          }
        }
      }
    }
  } catch (error) {
    console.log(`   Error during detection: ${error.message}`);
    return true;
  }
  return false;
};

const hasFormatting = detectFileFormatting(workbook);
console.log(`📊 Overall formatting detected: ${hasFormatting ? '✅ YES' : '❌ NO'}`);

// Try different reading methods
console.log('\n🧪 TESTING DIFFERENT READ METHODS:');

// Method 1: Minimal options
console.log('Method 1: Minimal options');
const wb1 = XLSX.read(fileData, { type: "buffer" });
console.log(`   Sheets: ${wb1.SheetNames.join(', ')}`);
console.log(`   First sheet formatting: ${wb1.Sheets[wb1.SheetNames[0]]['!cols'] ? 'YES' : 'NO'}`);

// Method 2: Only cellStyles
console.log('Method 2: Only cellStyles');
const wb2 = XLSX.read(fileData, { type: "buffer", cellStyles: true });
console.log(`   Sheets: ${wb2.SheetNames.join(', ')}`);
console.log(`   First sheet formatting: ${wb2.Sheets[wb2.SheetNames[0]]['!cols'] ? 'YES' : 'NO'}`);

// Method 3: All options
console.log('Method 3: All options (current)');
console.log(`   Sheets: ${workbook.SheetNames.join(', ')}`);
console.log(`   First sheet formatting: ${workbook.Sheets[workbook.SheetNames[0]]['!cols'] ? 'YES' : 'NO'}`);

// Raw file analysis
console.log('\n📁 RAW FILE ANALYSIS:');
const fileSize = fs.statSync('DATA.xlsx').size;
console.log(`   File size: ${(fileSize / 1024).toFixed(1)} KB`);

// Try to detect if it's a real Excel file
const fileHeader = fileData.slice(0, 4);
console.log(`   File header: ${fileHeader.toString('hex')} (${fileHeader.toString()})`);

// Check if it's a ZIP file (modern Excel format)
const isZip = fileHeader[0] === 0x50 && fileHeader[1] === 0x4B;
console.log(`   Is ZIP format (modern Excel): ${isZip ? 'YES' : 'NO'}`);

console.log('\n🎯 TỔNG KẾT:');
console.log('Nếu bạn thấy file có background colors và formatting trong Excel');
console.log('nhưng script này không detect được, có thể:');
console.log('1. ❓ File sử dụng conditional formatting (không được XLSX library support)');
console.log('2. ❓ File có theme-based colors (không được detect đúng)'); 
console.log('3. ❓ File có custom styling không standard');
console.log('4. ❓ Cần method đọc file khác');
console.log('\n💡 Hãy mở file DATA.xlsx trong Excel và screenshot để tôi hiểu rõ hơn!');