// DEBUG THỰC SỰ: Tại sao sheet 1 bị mất và formatting bị mất
const XLSX = require('xlsx');
const fs = require('fs');

console.log('🐛 DEBUGGING REAL ISSUES\n');

// Đọc file gốc
console.log('📂 Analyzing original DATA.xlsx...');
const originalData = fs.readFileSync('DATA.xlsx');
const originalWb = XLSX.read(originalData, {
  type: "buffer",
  cellStyles: true,
  cellFormula: true,
  cellDates: true,
  cellNF: true,
  bookVBA: true
});

console.log('ORIGINAL FILE:');
console.log(`   Sheets: ${originalWb.SheetNames.join(', ')}`);

// Check sheet 1 formatting
const sheet1Name = originalWb.SheetNames[0];
const originalSheet1 = originalWb.Sheets[sheet1Name];
console.log(`\nORIGINAL SHEET 1 "${sheet1Name}":`);
console.log(`   Range: ${originalSheet1['!ref']}`);
console.log(`   A1 value: "${originalSheet1['A1']?.v}"`);
console.log(`   A1 style: ${JSON.stringify(originalSheet1['A1']?.s || 'NO STYLE')}`);
console.log(`   Columns: ${originalSheet1['!cols'] ? originalSheet1['!cols'].length : 'NO'}`);
console.log(`   Merges: ${originalSheet1['!merges'] ? originalSheet1['!merges'].length : 'NO'}`);

// Check sheet 2 formatting  
const sheet2Name = originalWb.SheetNames[1];
const originalSheet2 = originalWb.Sheets[sheet2Name];
console.log(`\nORIGINAL SHEET 2 "${sheet2Name}":`);
console.log(`   Range: ${originalSheet2['!ref']}`);
console.log(`   A1 value: "${originalSheet2['A1']?.v}"`);
console.log(`   A1 style: ${JSON.stringify(originalSheet2['A1']?.s || 'NO STYLE')}`);
console.log(`   Columns: ${originalSheet2['!cols'] ? originalSheet2['!cols'].length : 'NO'}`);

// Đọc file modified
console.log('\n📂 Analyzing attendee_report_updated_FINAL.xlsx...');

if (!fs.existsSync('attendee_report_updated_FINAL.xlsx')) {
  console.log('❌ Modified file not found!');
  process.exit(1);
}

const modifiedData = fs.readFileSync('attendee_report_updated_FINAL.xlsx');
const modifiedWb = XLSX.read(modifiedData, {
  type: "buffer", 
  cellStyles: true,
  cellFormula: true,
  cellDates: true,
  cellNF: true,
  bookVBA: true
});

console.log('MODIFIED FILE:');
console.log(`   Sheets: ${modifiedWb.SheetNames.join(', ')}`);

if (modifiedWb.SheetNames.length < originalWb.SheetNames.length) {
  console.log('❌ MISSING SHEETS DETECTED!');
  console.log(`   Original had: ${originalWb.SheetNames.length} sheets`);
  console.log(`   Modified has: ${modifiedWb.SheetNames.length} sheets`);
}

// Check if sheet 1 exists in modified
if (modifiedWb.Sheets[sheet1Name]) {
  const modifiedSheet1 = modifiedWb.Sheets[sheet1Name];
  console.log(`\nMODIFIED SHEET 1 "${sheet1Name}": EXISTS`);
  console.log(`   Range: ${modifiedSheet1['!ref']}`);
  console.log(`   A1 value: "${modifiedSheet1['A1']?.v}"`);
  console.log(`   A1 style: ${JSON.stringify(modifiedSheet1['A1']?.s || 'NO STYLE')}`);
  console.log(`   Columns: ${modifiedSheet1['!cols'] ? modifiedSheet1['!cols'].length : 'NO'}`);
  console.log(`   Merges: ${modifiedSheet1['!merges'] ? modifiedSheet1['!merges'].length : 'NO'}`);
} else {
  console.log(`\n❌ SHEET 1 "${sheet1Name}": MISSING FROM MODIFIED FILE!`);
}

// Check sheet 2 in modified
if (modifiedWb.Sheets[sheet2Name]) {
  const modifiedSheet2 = modifiedWb.Sheets[sheet2Name];
  console.log(`\nMODIFIED SHEET 2 "${sheet2Name}": EXISTS`);
  console.log(`   Range: ${modifiedSheet2['!ref']}`);
  console.log(`   A1 value: "${modifiedSheet2['A1']?.v}"`);
  console.log(`   A1 style: ${JSON.stringify(modifiedSheet2['A1']?.s || 'NO STYLE')}`);
  console.log(`   Columns: ${modifiedSheet2['!cols'] ? modifiedSheet2['!cols'].length : 'NO'}`);
} else {
  console.log(`\n❌ SHEET 2 "${sheet2Name}": MISSING FROM MODIFIED FILE!`);
}

// Detailed A1 comparison
console.log('\n🔍 DETAILED A1 COMPARISON:');
console.log('ORIGINAL A1:');
if (originalSheet2['A1']) {
  console.log(`   Value: "${originalSheet2['A1'].v}"`);
  console.log(`   Type: ${originalSheet2['A1'].t}`);
  console.log(`   Style: ${JSON.stringify(originalSheet2['A1'].s, null, 2)}`);
} else {
  console.log('   A1 NOT FOUND');
}

if (modifiedWb.Sheets[sheet2Name] && modifiedWb.Sheets[sheet2Name]['A1']) {
  console.log('MODIFIED A1:');
  const modA1 = modifiedWb.Sheets[sheet2Name]['A1'];
  console.log(`   Value: "${modA1.v}"`);
  console.log(`   Type: ${modA1.t}`);
  console.log(`   Style: ${JSON.stringify(modA1.s, null, 2)}`);
} else {
  console.log('MODIFIED A1: NOT FOUND');
}

// Identify the real issues
console.log('\n🎯 IDENTIFIED ISSUES:');

const issues = [];

if (modifiedWb.SheetNames.length < originalWb.SheetNames.length) {
  issues.push('❌ Missing sheets in output');
}

if (!modifiedWb.Sheets[sheet1Name]) {
  issues.push(`❌ Sheet "${sheet1Name}" completely missing`);
}

if (originalSheet2['A1']?.s && (!modifiedWb.Sheets[sheet2Name]?.['A1']?.s || 
    JSON.stringify(originalSheet2['A1'].s) !== JSON.stringify(modifiedWb.Sheets[sheet2Name]['A1'].s))) {
  issues.push('❌ A1 formatting lost (border, bold, etc.)');
}

if (issues.length === 0) {
  console.log('✅ No issues found - might be a reading problem');
} else {
  issues.forEach(issue => console.log(`   ${issue}`));
}

// Test simple fix
console.log('\n🔧 TESTING SIMPLE FIX...');

try {
  // Create a proper copy that preserves ALL sheets
  const testWb = XLSX.utils.book_new();
  
  // Copy ALL sheets from original
  originalWb.SheetNames.forEach(sheetName => {
    console.log(`   Copying sheet: ${sheetName}`);
    
    const originalSheet = originalWb.Sheets[sheetName];
    const copiedSheet = {};
    
    // Deep copy every property
    Object.keys(originalSheet).forEach(key => {
      const value = originalSheet[key];
      if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
        copiedSheet[key] = JSON.parse(JSON.stringify(value));
      } else if (Array.isArray(value)) {
        copiedSheet[key] = value.map(item => 
          typeof item === 'object' && item !== null ? JSON.parse(JSON.stringify(item)) : item
        );
      } else {
        copiedSheet[key] = value;
      }
    });
    
    testWb.Sheets[sheetName] = copiedSheet;
  });
  
  testWb.SheetNames = [...originalWb.SheetNames];
  
  // Only modify sheet 2, leave sheet 1 untouched
  if (testWb.Sheets[sheet2Name]) {
    const sheet2 = testWb.Sheets[sheet2Name];
    const currentRange = XLSX.utils.decode_range(sheet2['!ref'] || 'A1:J1000');
    const newColIndex = currentRange.e.c + 1;
    const newColLetter = XLSX.utils.encode_col(newColIndex);
    
    // Add check-in header with PRESERVED formatting
    const headerAddr = `${newColLetter}1`;
    const originalA1Style = sheet2['A1']?.s;
    
    sheet2[headerAddr] = {
      v: 'CHECK-IN TIME',
      t: 's',
      ...(originalA1Style && { s: JSON.parse(JSON.stringify(originalA1Style)) })
    };
    
    // Add sample data
    sheet2[`${newColLetter}2`] = { v: '10:30 AM', t: 's' };
    sheet2[`${newColLetter}3`] = { v: '11:15 AM', t: 's' };
    
    // Update range
    sheet2['!ref'] = XLSX.utils.encode_range({
      s: currentRange.s,
      e: { r: currentRange.e.r, c: newColIndex }
    });
    
    // Update columns
    if (sheet2['!cols']) {
      sheet2['!cols'] = [...sheet2['!cols']];
      sheet2['!cols'][newColIndex] = { width: 15 };
    }
    
    console.log(`   Modified sheet 2: added column ${newColLetter}`);
  }
  
  // Write with ALL formatting options
  const fixedBuffer = XLSX.write(testWb, {
    bookType: 'xlsx',
    type: 'array',
    cellStyles: true,
    cellDates: true,
    cellFormula: true,
    cellNF: true,
    bookVBA: true,
    compression: false // Try without compression
  });
  
  fs.writeFileSync('DATA_FIXED.xlsx', Buffer.from(fixedBuffer));
  console.log('✅ Created DATA_FIXED.xlsx');
  
  // Verify the fix
  const verifyWb = XLSX.read(fixedBuffer, {
    type: "array",
    cellStyles: true,
    cellFormula: true,
    cellDates: true,
    cellNF: true,
    bookVBA: true
  });
  
  console.log('\n📊 VERIFICATION OF FIX:');
  console.log(`   Sheets: ${verifyWb.SheetNames.join(', ')}`);
  console.log(`   Sheet count: ${verifyWb.SheetNames.length} (should be ${originalWb.SheetNames.length})`);
  
  if (verifyWb.Sheets[sheet1Name]) {
    console.log(`   ✅ Sheet 1 "${sheet1Name}": PRESERVED`);
  } else {
    console.log(`   ❌ Sheet 1 "${sheet1Name}": STILL MISSING`);
  }
  
  if (verifyWb.Sheets[sheet2Name] && verifyWb.Sheets[sheet2Name]['A1']) {
    const fixedA1 = verifyWb.Sheets[sheet2Name]['A1'];
    console.log(`   A1 value: "${fixedA1.v}"`);
    console.log(`   A1 style preserved: ${fixedA1.s ? 'YES' : 'NO'}`);
    if (fixedA1.s) {
      console.log(`   A1 style: ${JSON.stringify(fixedA1.s)}`);
    }
  }
  
} catch (error) {
  console.log(`❌ Fix failed: ${error.message}`);
}

console.log('\n🎯 SUMMARY:');
console.log('Bạn đúng rồi - có bug thực sự trong code!');
console.log('📁 Hãy thử mở file DATA_FIXED.xlsx để xem có khác biệt không');
console.log('Nếu vẫn lỗi, tôi sẽ fix deeper vào code handling.');