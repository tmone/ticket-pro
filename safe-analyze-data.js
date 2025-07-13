// Phân tích an toàn file DATA.xlsx
const XLSX = require('xlsx');
const fs = require('fs');

console.log('🔍 PHÂN TÍCH AN TOÀN FILE DATA.xlsx\n');

if (!fs.existsSync('DATA.xlsx')) {
  console.log('❌ File DATA.xlsx không tồn tại!');
  process.exit(1);
}

try {
  const fileData = fs.readFileSync('DATA.xlsx');
  
  // Basic file info
  console.log('📁 FILE INFO:');
  const fileSize = fs.statSync('DATA.xlsx').size;
  console.log(`   Size: ${(fileSize / 1024).toFixed(1)} KB`);
  
  const fileHeader = fileData.slice(0, 8);
  console.log(`   Header: ${fileHeader.toString('hex')}`);
  
  const isZip = fileHeader[0] === 0x50 && fileHeader[1] === 0x4B;
  console.log(`   Format: ${isZip ? 'Modern Excel (ZIP-based)' : 'Legacy Excel'}`);
  
  // Try different read methods
  console.log('\n🧪 TESTING READ METHODS:');
  
  // Method 1: Basic read
  console.log('\n📖 Method 1: Basic read');
  const wb1 = XLSX.read(fileData, { type: "buffer" });
  console.log(`   Success: YES`);
  console.log(`   Sheets: ${wb1.SheetNames.length} (${wb1.SheetNames.join(', ')})`);
  
  // Method 2: With cellStyles
  console.log('\n🎨 Method 2: With cellStyles');
  const wb2 = XLSX.read(fileData, { 
    type: "buffer", 
    cellStyles: true 
  });
  console.log(`   Success: YES`);
  console.log(`   Sheets: ${wb2.SheetNames.length} (${wb2.SheetNames.join(', ')})`);
  
  // Method 3: Maximum options
  console.log('\n🔧 Method 3: Maximum options');
  const wb3 = XLSX.read(fileData, {
    type: "buffer",
    cellStyles: true,
    cellFormula: true,
    cellDates: true,
    cellNF: true,
    sheetStubs: true,
    bookVBA: true
  });
  console.log(`   Success: YES`);
  console.log(`   Sheets: ${wb3.SheetNames.length} (${wb3.SheetNames.join(', ')})`);
  
  // Use the best workbook for analysis
  const workbook = wb3;
  
  // Analyze each sheet safely
  console.log('\n📊 SHEET ANALYSIS:');
  
  workbook.SheetNames.forEach((sheetName, index) => {
    console.log(`\n🗂️  SHEET ${index + 1}: "${sheetName}"`);
    
    try {
      const sheet = workbook.Sheets[sheetName];
      
      if (!sheet) {
        console.log('   ❌ Sheet is null/undefined');
        return;
      }
      
      // Basic info
      console.log(`   📐 Range: ${sheet['!ref'] || 'NO RANGE'}`);
      
      // Count different elements
      let totalCells = 0;
      let styledCells = 0;
      let formulaCells = 0;
      let commentCells = 0;
      let hyperlinkCells = 0;
      
      const sampleCells = [];
      const styledCellsSample = [];
      
      Object.keys(sheet).forEach(key => {
        if (!key.startsWith('!')) {
          totalCells++;
          const cell = sheet[key];
          
          // Sample first few cells
          if (sampleCells.length < 10) {
            sampleCells.push({
              addr: key,
              value: cell.v,
              type: cell.t,
              hasStyle: !!cell.s,
              hasFormula: !!cell.f
            });
          }
          
          if (cell.s) {
            styledCells++;
            if (styledCellsSample.length < 5) {
              styledCellsSample.push({
                addr: key,
                value: cell.v,
                style: cell.s
              });
            }
          }
          if (cell.f) formulaCells++;
          if (cell.c) commentCells++;
          if (cell.l) hyperlinkCells++;
        }
      });
      
      // Sheet-level formatting
      console.log('   🎨 Sheet formatting:');
      console.log(`      Columns (!cols): ${sheet['!cols'] ? `YES (${sheet['!cols'].length})` : 'NO'}`);
      console.log(`      Rows (!rows): ${sheet['!rows'] ? `YES (${sheet['!rows'].length})` : 'NO'}`);
      console.log(`      Merges (!merges): ${sheet['!merges'] ? `YES (${sheet['!merges'].length})` : 'NO'}`);
      console.log(`      Auto filter: ${sheet['!autofilter'] ? 'YES' : 'NO'}`);
      console.log(`      Protection: ${sheet['!protect'] ? 'YES' : 'NO'}`);
      
      // Cell statistics
      console.log('   📊 Cell statistics:');
      console.log(`      Total cells: ${totalCells}`);
      console.log(`      Styled cells: ${styledCells}`);
      console.log(`      Formula cells: ${formulaCells}`);
      console.log(`      Comment cells: ${commentCells}`);
      console.log(`      Hyperlink cells: ${hyperlinkCells}`);
      
      // Sample data
      console.log('   📋 Sample cells:');
      sampleCells.forEach(cell => {
        console.log(`      ${cell.addr}: "${cell.value}" (type: ${cell.type}, styled: ${cell.hasStyle})`);
      });
      
      // Sample styled cells
      if (styledCellsSample.length > 0) {
        console.log('   🎨 Sample styled cells:');
        styledCellsSample.forEach(cell => {
          const styleStr = JSON.stringify(cell.style);
          const truncated = styleStr.length > 100 ? styleStr.substring(0, 100) + '...' : styleStr;
          console.log(`      ${cell.addr}: "${cell.value}" style: ${truncated}`);
        });
      }
      
      // Detailed column info
      if (sheet['!cols'] && sheet['!cols'].length > 0) {
        console.log('   📏 Column details (first 5):');
        sheet['!cols'].slice(0, 5).forEach((col, i) => {
          if (col) {
            console.log(`      Col ${i}: ${JSON.stringify(col)}`);
          }
        });
      }
      
      // Detailed merge info
      if (sheet['!merges'] && sheet['!merges'].length > 0) {
        console.log('   🔗 Merge details (first 3):');
        sheet['!merges'].slice(0, 3).forEach((merge, i) => {
          const start = XLSX.utils.encode_cell(merge.s);
          const end = XLSX.utils.encode_cell(merge.e);
          console.log(`      ${i + 1}: ${start} to ${end}`);
        });
      }
      
    } catch (error) {
      console.log(`   ❌ Error analyzing sheet: ${error.message}`);
    }
  });
  
  // Overall detection
  console.log('\n🔍 FORMATTING DETECTION:');
  
  let hasAnyFormatting = false;
  let formatDetails = [];
  
  workbook.SheetNames.forEach(sheetName => {
    try {
      const sheet = workbook.Sheets[sheetName];
      
      if (sheet['!cols']) {
        hasAnyFormatting = true;
        formatDetails.push(`${sheetName}: Column widths`);
      }
      if (sheet['!rows']) {
        hasAnyFormatting = true;
        formatDetails.push(`${sheetName}: Row heights`);
      }
      if (sheet['!merges']) {
        hasAnyFormatting = true;
        formatDetails.push(`${sheetName}: Merged cells`);
      }
      
      // Check for styled cells
      let hasStyledCells = false;
      Object.keys(sheet).forEach(key => {
        if (!key.startsWith('!') && sheet[key].s) {
          hasStyledCells = true;
        }
      });
      
      if (hasStyledCells) {
        hasAnyFormatting = true;
        formatDetails.push(`${sheetName}: Cell styles`);
      }
      
    } catch (error) {
      console.log(`Error checking ${sheetName}: ${error.message}`);
    }
  });
  
  console.log(`📊 Overall formatting found: ${hasAnyFormatting ? '✅ YES' : '❌ NO'}`);
  if (formatDetails.length > 0) {
    console.log('📋 Details:');
    formatDetails.forEach(detail => console.log(`   • ${detail}`));
  }
  
  // Test modification
  console.log('\n✏️  TESTING MODIFICATION:');
  
  try {
    // Create a copy
    const testWb = JSON.parse(JSON.stringify(workbook));
    
    // Try to modify first sheet
    const firstSheetName = workbook.SheetNames[0];
    const firstSheet = testWb.Sheets[firstSheetName];
    
    // Find a cell to modify
    let cellToModify = null;
    Object.keys(firstSheet).forEach(key => {
      if (!key.startsWith('!') && !cellToModify) {
        cellToModify = key;
      }
    });
    
    if (cellToModify) {
      const originalValue = firstSheet[cellToModify].v;
      firstSheet[cellToModify].v = `${originalValue} [TEST]`;
      
      console.log(`✅ Modified ${cellToModify}: "${originalValue}" → "${firstSheet[cellToModify].v}"`);
      
      // Try to write
      const testBuffer = XLSX.write(testWb, {
        bookType: 'xlsx',
        type: 'array',
        cellStyles: true
      });
      
      fs.writeFileSync('DATA_TEST_MODIFY.xlsx', Buffer.from(testBuffer));
      console.log('✅ Test modification file created: DATA_TEST_MODIFY.xlsx');
      
      // Verify
      const verifyWb = XLSX.read(testBuffer, {
        type: "array",
        cellStyles: true
      });
      
      const verifySheet = verifyWb.Sheets[firstSheetName];
      const verifiedValue = verifySheet[cellToModify].v;
      
      console.log(`🔍 Verification: ${verifiedValue.includes('[TEST]') ? '✅ SUCCESS' : '❌ FAILED'}`);
      
    } else {
      console.log('❌ No cell found to modify');
    }
    
  } catch (error) {
    console.log(`❌ Modification test failed: ${error.message}`);
  }
  
} catch (error) {
  console.log(`❌ Overall error: ${error.message}`);
  console.log('Stack:', error.stack);
}

console.log('\n🎯 SUMMARY:');
console.log('Nếu bạn vẫn thấy script không detect được formatting');
console.log('mặc dù Excel hiển thị colors/formatting, có thể do:');
console.log('1. 📱 Conditional formatting (dynamic, không lưu trong cell)');
console.log('2. 🎨 Theme-based colors (colors từ theme, không explicit)');
console.log('3. 📊 Table formatting (Excel tables có formatting riêng)');
console.log('4. 🔧 Worksheet protection/hidden formatting');
console.log('\n💡 Hãy thử mở file DATA_TEST_MODIFY.xlsx để so sánh!');