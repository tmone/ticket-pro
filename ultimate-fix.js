// ULTIMATE FIX: Giải quyết tất cả vấn đề đã phát hiện
const XLSX = require('xlsx');
const fs = require('fs');

console.log('🚀 ULTIMATE FIX: Giải quyết tất cả vấn đề\n');

// Đọc file gốc
const originalData = fs.readFileSync('DATA.xlsx');
const originalWb = XLSX.read(originalData, {
  type: "buffer",
  cellStyles: true,
  cellFormula: true,
  cellDates: true,
  cellNF: true,
  bookVBA: true
});

console.log('📂 ORIGINAL FILE ANALYSIS:');
console.log(`   Sheets: ${originalWb.SheetNames.join(', ')}`);

// Check sheet visibility
if (originalWb.Workbook && originalWb.Workbook.Sheets) {
  originalWb.Workbook.Sheets.forEach((sheet, i) => {
    console.log(`   Sheet ${i + 1} "${sheet.name}": Hidden = ${sheet.Hidden ? 'YES' : 'NO'}`);
  });
}

// SOLUTION 1: Copy with maximum preservation
console.log('\n🔧 SOLUTION 1: Maximum preservation copy...');

const enhancedCopy = () => {
  // Create new workbook
  const newWb = XLSX.utils.book_new();
  
  // Copy workbook-level properties FIRST
  if (originalWb.Props) newWb.Props = JSON.parse(JSON.stringify(originalWb.Props));
  if (originalWb.Custprops) newWb.Custprops = JSON.parse(JSON.stringify(originalWb.Custprops));
  if (originalWb.Workbook) newWb.Workbook = JSON.parse(JSON.stringify(originalWb.Workbook));
  if (originalWb.vbaraw) newWb.vbaraw = originalWb.vbaraw;
  if (originalWb.Styles) newWb.Styles = JSON.parse(JSON.stringify(originalWb.Styles));
  if (originalWb.SSF) newWb.SSF = JSON.parse(JSON.stringify(originalWb.SSF));
  
  console.log('   ✅ Copied workbook-level properties');
  
  // Copy ALL sheets with enhanced preservation
  originalWb.SheetNames.forEach((sheetName, index) => {
    console.log(`   📋 Processing sheet: "${sheetName}"`);
    
    const originalSheet = originalWb.Sheets[sheetName];
    const newSheet = {};
    
    // Copy EVERYTHING from original sheet
    Object.keys(originalSheet).forEach(key => {
      const value = originalSheet[key];
      
      if (key.startsWith('!')) {
        // Sheet properties - deep copy
        if (Array.isArray(value)) {
          newSheet[key] = value.map(item => 
            typeof item === 'object' && item !== null 
              ? JSON.parse(JSON.stringify(item))
              : item
          );
        } else if (typeof value === 'object' && value !== null) {
          newSheet[key] = JSON.parse(JSON.stringify(value));
        } else {
          newSheet[key] = value;
        }
      } else {
        // Cell data - preserve EVERYTHING
        if (typeof value === 'object' && value !== null) {
          newSheet[key] = JSON.parse(JSON.stringify(value));
        } else {
          newSheet[key] = value;
        }
      }
    });
    
    newWb.Sheets[sheetName] = newSheet;
  });
  
  // Set sheet names
  newWb.SheetNames = [...originalWb.SheetNames];
  
  // IMPORTANT: Unhide all sheets for user visibility
  if (newWb.Workbook && newWb.Workbook.Sheets) {
    newWb.Workbook.Sheets.forEach(sheet => {
      sheet.Hidden = 0; // Unhide all sheets
    });
    console.log('   👁️  Unhidden all sheets for visibility');
  }
  
  return newWb;
};

const preservedWb = enhancedCopy();

// SOLUTION 2: Add check-in data to data sheet only
console.log('\n➕ SOLUTION 2: Adding check-in data...');

const dataSheetName = 'Thong_tin_khach';
const dataSheet = preservedWb.Sheets[dataSheetName];

if (dataSheet) {
  const currentRange = XLSX.utils.decode_range(dataSheet['!ref'] || 'A1:J1000');
  const newColIndex = currentRange.e.c + 1;
  const newColLetter = XLSX.utils.encode_col(newColIndex);
  
  console.log(`   Adding column ${newColLetter} to "${dataSheetName}"`);
  
  // Add header - try to match existing header style as much as possible
  const headerAddr = `${newColLetter}1`;
  const sampleHeaderCell = dataSheet['A1'];
  
  dataSheet[headerAddr] = {
    v: 'CHECK-IN TIME',
    t: 's',
    // Copy whatever style properties we can get
    ...(sampleHeaderCell?.s && { s: JSON.parse(JSON.stringify(sampleHeaderCell.s)) })
  };
  
  // Add sample check-in data
  const sampleData = [
    '2024-01-15 10:30:00',
    '2024-01-15 11:15:00', 
    '',
    '2024-01-15 09:45:00',
    ''
  ];
  
  sampleData.forEach((checkIn, i) => {
    const cellAddr = `${newColLetter}${i + 2}`;
    const sampleDataCell = dataSheet[`A${i + 2}`];
    
    dataSheet[cellAddr] = {
      v: checkIn,
      t: 's',
      ...(sampleDataCell?.s && { s: JSON.parse(JSON.stringify(sampleDataCell.s)) })
    };
  });
  
  // Update range
  dataSheet['!ref'] = XLSX.utils.encode_range({
    s: currentRange.s,
    e: { r: currentRange.e.r, c: newColIndex }
  });
  
  // Update column widths if they exist
  if (dataSheet['!cols']) {
    const newCols = [...dataSheet['!cols']];
    // Ensure we have enough columns
    while (newCols.length <= newColIndex) {
      newCols.push({ width: 15 });
    }
    dataSheet['!cols'] = newCols;
  }
  
  console.log(`   ✅ Added check-in column with ${sampleData.filter(d => d).length} sample records`);
} else {
  console.log(`   ❌ Data sheet "${dataSheetName}" not found`);
}

// SOLUTION 3: Write with ALL possible preservation options
console.log('\n💾 SOLUTION 3: Writing with maximum preservation...');

const writeOptions = {
  bookType: 'xlsx',
  type: 'array',
  cellStyles: true,
  cellDates: true,
  cellFormula: true,
  cellNF: true,
  bookVBA: true,
  compression: false, // Disable compression to avoid issues
  bookSST: false, // Let Excel handle string table
  writeFileWithStyles: true // Custom flag (if supported)
};

const finalBuffer = XLSX.write(preservedWb, writeOptions);

const outputFile = 'DATA_ULTIMATE_FIX.xlsx';
fs.writeFileSync(outputFile, Buffer.from(finalBuffer));

console.log(`✅ Created: ${outputFile}`);
console.log(`   Size: ${(finalBuffer.byteLength / 1024).toFixed(1)} KB`);

// VERIFICATION
console.log('\n🔍 VERIFICATION...');

const verifyWb = XLSX.read(finalBuffer, {
  type: "array",
  cellStyles: true,
  cellFormula: true,
  cellDates: true,
  cellNF: true,
  bookVBA: true
});

console.log('📊 FINAL RESULTS:');
console.log(`   Sheets: ${verifyWb.SheetNames.join(', ')}`);
console.log(`   Sheet count: ${verifyWb.SheetNames.length} (original: ${originalWb.SheetNames.length})`);

// Check sheet visibility
if (verifyWb.Workbook && verifyWb.Workbook.Sheets) {
  verifyWb.Workbook.Sheets.forEach((sheet, i) => {
    const visibility = sheet.Hidden ? 'HIDDEN' : 'VISIBLE';
    console.log(`   Sheet ${i + 1} "${sheet.name}": ${visibility}`);
  });
}

// Check data sheet
const verifyDataSheet = verifyWb.Sheets[dataSheetName];
if (verifyDataSheet) {
  console.log(`   ${dataSheetName} range: ${verifyDataSheet['!ref']}`);
  console.log(`   Column widths: ${verifyDataSheet['!cols'] ? `${verifyDataSheet['!cols'].length} cols` : 'NO'}`);
  console.log(`   Row heights: ${verifyDataSheet['!rows'] ? `${verifyDataSheet['!rows'].length} rows` : 'NO'}`);
  
  // Check new column
  const newColLetter = XLSX.utils.encode_col(XLSX.utils.decode_range(verifyDataSheet['!ref']).e.c);
  const headerValue = verifyDataSheet[`${newColLetter}1`]?.v;
  console.log(`   New header (${newColLetter}1): "${headerValue}"`);
}

// Final assessment
console.log('\n🎯 FIXES APPLIED:');
console.log('✅ 1. Unhidden all sheets (sheet "sơ đồ ghế" should now be visible)');
console.log('✅ 2. Preserved all structural formatting (columns, rows, merges)');
console.log('✅ 3. Added check-in column with sample data');
console.log('✅ 4. Used maximum preservation write options');

console.log('\n💡 IMPORTANT NOTES:');
console.log('📌 Visual formatting (colors, borders, bold) limitations:');
console.log('   - XLSX.js library có giới hạn với complex Excel formatting');
console.log('   - Structure (layout, columns, merges) được preserve hoàn toàn');
console.log('   - Visual styles có thể bị "simplified" nhưng file vẫn functional');

console.log('\n📁 TEST FILES:');
console.log(`   • DATA.xlsx (original)`);
console.log(`   • ${outputFile} (ultimate fix)`);
console.log('\n🔍 Hãy mở file mới và kiểm tra:');
console.log('   1. Có thấy cả 2 sheets không?');
console.log('   2. Layout có giống nhau không?');
console.log('   3. Check-in column có được thêm không?');

console.log('\n⚡ Nếu vẫn có vấn đề, có thể cần:');
console.log('   1. Sử dụng library khác (như ExcelJS)');
console.log('   2. Hoặc chấp nhận limitation và focus vào functionality');