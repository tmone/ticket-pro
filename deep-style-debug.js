// Deep debug: Tại sao style chỉ là {"patternType":"none"}
const XLSX = require('xlsx');
const fs = require('fs');

console.log('🔍 DEEP STYLE DEBUGGING\n');

// Test 1: Thử đọc với tất cả các options khác nhau
console.log('TEST 1: Different read options...');

const fileData = fs.readFileSync('DATA.xlsx');

const readMethods = [
  { name: 'Basic', options: { type: "buffer" } },
  { name: 'CellStyles Only', options: { type: "buffer", cellStyles: true } },
  { name: 'All Options', options: { 
    type: "buffer", 
    cellStyles: true, 
    cellFormula: true, 
    cellDates: true, 
    cellNF: true,
    bookVBA: true,
    sheetStubs: true
  }},
  { name: 'Raw Mode', options: { 
    type: "buffer", 
    cellStyles: true,
    raw: true 
  }},
  { name: 'Dense Mode', options: { 
    type: "buffer", 
    cellStyles: true,
    dense: true 
  }}
];

readMethods.forEach(method => {
  try {
    console.log(`\n${method.name}:`);
    const wb = XLSX.read(fileData, method.options);
    const sheet = wb.Sheets['Thong_tin_khach'];
    const a1 = sheet['A1'];
    
    console.log(`   A1 value: "${a1?.v}"`);
    console.log(`   A1 type: ${a1?.t}`);
    console.log(`   A1 style: ${JSON.stringify(a1?.s || 'NO STYLE')}`);
    
    // Check more cells
    const b1 = sheet['B1'];
    if (b1) {
      console.log(`   B1 value: "${b1.v}" style: ${JSON.stringify(b1.s || 'NO STYLE')}`);
    }
  } catch (error) {
    console.log(`   ERROR: ${error.message}`);
  }
});

// Test 2: Analyze file structure directly
console.log('\n\nTEST 2: File structure analysis...');

// Read as zip and check internal files
const AdmZip = require('fs').existsSync('node_modules/adm-zip') ? require('adm-zip') : null;

if (!AdmZip) {
  console.log('adm-zip not available, trying manual analysis...');
  
  // Check if it's a valid zip
  const header = fileData.slice(0, 4);
  console.log(`File header: ${header.toString('hex')} (${header.toString()})`);
  
  if (header[0] === 0x50 && header[1] === 0x4B) {
    console.log('✅ Valid ZIP format (modern Excel)');
    
    // Try to read with different type specifications
    const typeTests = ['buffer', 'array', 'base64'];
    
    typeTests.forEach(type => {
      try {
        console.log(`\nTesting type: ${type}`);
        let data = fileData;
        if (type === 'array') data = Array.from(fileData);
        if (type === 'base64') data = fileData.toString('base64');
        
        const wb = XLSX.read(data, { type, cellStyles: true });
        const sheet = wb.Sheets['Thong_tin_khach'];
        console.log(`   Success! A1 style: ${JSON.stringify(sheet['A1']?.s || 'NO STYLE')}`);
      } catch (error) {
        console.log(`   Failed: ${error.message}`);
      }
    });
  } else {
    console.log('❌ Not a ZIP format - might be legacy Excel');
  }
}

// Test 3: Create a simple test file with known formatting
console.log('\n\nTEST 3: Create test file with formatting...');

const testWb = XLSX.utils.book_new();
const testData = [
  ['Header 1', 'Header 2'],
  ['Data A', 'Data B']
];

const testWs = XLSX.utils.aoa_to_sheet(testData);

// Add explicit formatting
testWs['A1'].s = {
  font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
  fill: { fgColor: { rgb: 'FF0000' }, patternType: 'solid' },
  border: {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } }
  },
  alignment: { horizontal: 'center', vertical: 'center' }
};

testWs['B1'].s = {
  font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
  fill: { fgColor: { rgb: '0000FF' }, patternType: 'solid' },
  border: {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } }
  }
};

testWs['!cols'] = [{ width: 15 }, { width: 20 }];

XLSX.utils.book_append_sheet(testWb, testWs, 'TestFormatting');

console.log('Created test sheet with formatting:');
console.log(`   A1 style: ${JSON.stringify(testWs['A1'].s)}`);

// Write and read back
const testBuffer = XLSX.write(testWb, {
  type: 'array',
  bookType: 'xlsx',
  cellStyles: true
});

console.log('Writing test file...');
fs.writeFileSync('FORMAT_TEST.xlsx', Buffer.from(testBuffer));

console.log('Reading back test file...');
const readBackWb = XLSX.read(testBuffer, {
  type: 'array',
  cellStyles: true
});

const readBackWs = readBackWb.Sheets['TestFormatting'];
console.log(`   Read back A1 style: ${JSON.stringify(readBackWs['A1']?.s || 'NO STYLE')}`);

// Test 4: Check if there are hidden styles in the original
console.log('\n\nTEST 4: Check for hidden/complex styles...');

const wb = XLSX.read(fileData, { type: "buffer", cellStyles: true });
const sheet = wb.Sheets['Thong_tin_khach'];

// Check first 10 cells for any styles
console.log('Checking first 10 cells for styles:');
let stylesFound = 0;
let cellsChecked = 0;

Object.keys(sheet).forEach(key => {
  if (!key.startsWith('!') && cellsChecked < 10) {
    cellsChecked++;
    const cell = sheet[key];
    
    console.log(`   ${key}: "${cell.v}" type:${cell.t}`);
    
    if (cell.s) {
      stylesFound++;
      const style = cell.s;
      console.log(`      Style keys: ${Object.keys(style).join(', ')}`);
      
      // Check each style property
      if (style.font) console.log(`      Font: ${JSON.stringify(style.font)}`);
      if (style.fill) console.log(`      Fill: ${JSON.stringify(style.fill)}`);
      if (style.border) console.log(`      Border: ${JSON.stringify(style.border)}`);
      if (style.alignment) console.log(`      Alignment: ${JSON.stringify(style.alignment)}`);
      if (style.numFmt) console.log(`      NumFmt: ${style.numFmt}`);
    } else {
      console.log(`      NO STYLE`);
    }
  }
});

console.log(`\nFound ${stylesFound} cells with styles out of ${cellsChecked} checked`);

// Test 5: Check workbook-level formatting info
console.log('\n\nTEST 5: Workbook-level formatting...');

if (wb.SSF) {
  console.log(`   Number formats available: ${Object.keys(wb.SSF).length}`);
}

if (wb.Workbook && wb.Workbook.Sheets) {
  console.log(`   Workbook sheet info: ${JSON.stringify(wb.Workbook.Sheets)}`);
}

// Look for style tables
console.log('   Checking for style-related properties...');
const wbKeys = Object.keys(wb);
wbKeys.forEach(key => {
  if (key.toLowerCase().includes('style') || key.toLowerCase().includes('format')) {
    console.log(`   Found: ${key} = ${typeof wb[key]}`);
  }
});

console.log('\n🎯 CONCLUSION:');
console.log('Nếu tất cả cells chỉ hiển thị {"patternType":"none"},');
console.log('có thể là:');
console.log('1. 📊 File sử dụng conditional formatting thay vì cell styles');
console.log('2. 🎨 File có theme-based formatting');
console.log('3. 📋 File có table formatting');
console.log('4. 🔧 XLSX library không đọc được style format này');
console.log('\n💡 Hãy thử mở FORMAT_TEST.xlsx để xem có style đúng không!');
console.log('Nếu FORMAT_TEST.xlsx có style, vấn đề là với file gốc.');
console.log('Nếu FORMAT_TEST.xlsx cũng mất style, vấn đề là với XLSX library.');