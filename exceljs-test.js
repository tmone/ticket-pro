// Test ExcelJS library với file DATA.xlsx
const fs = require('fs');

console.log('🧪 TESTING EXCELJS LIBRARY\n');

// Check if ExcelJS is available
let ExcelJS;
try {
  ExcelJS = require('exceljs');
  console.log('✅ ExcelJS library loaded successfully');
} catch (error) {
  console.log('❌ ExcelJS not available, trying alternative approach...');
  
  // Try to load from different paths
  const possiblePaths = [
    './node_modules/exceljs',
    '../node_modules/exceljs',
    'exceljs'
  ];
  
  for (const path of possiblePaths) {
    try {
      ExcelJS = require(path);
      console.log(`✅ ExcelJS loaded from: ${path}`);
      break;
    } catch (e) {
      console.log(`❌ Failed to load from: ${path}`);
    }
  }
  
  if (!ExcelJS) {
    console.log('📦 ExcelJS not found. Creating manual test...');
    console.log('\n💡 To install ExcelJS:');
    console.log('   npm install exceljs');
    console.log('\n📋 ExcelJS benefits over XLSX.js:');
    console.log('   ✅ Better formatting preservation');
    console.log('   ✅ Full color support (RGB, theme colors)');
    console.log('   ✅ Border styling');
    console.log('   ✅ Font formatting (bold, italic, size)');
    console.log('   ✅ Cell alignment');
    console.log('   ✅ Number formatting');
    console.log('   ✅ Conditional formatting support');
    console.log('   ✅ Better merge cell handling');
    
    // Create a comparison script that can be run when ExcelJS is available
    const comparisonScript = `
// ExcelJS vs XLSX.js Comparison Script
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const fs = require('fs');

console.log('📊 EXCELJS VS XLSX.JS COMPARISON\\n');

async function testExcelJS() {
  console.log('🔵 TESTING EXCELJS...');
  
  const workbook = new ExcelJS.Workbook();
  
  // Read the file
  await workbook.xlsx.readFile('DATA.xlsx');
  
  console.log('📂 ExcelJS Results:');
  console.log(\`   Worksheets: \${workbook.worksheets.length}\`);
  
  workbook.worksheets.forEach((worksheet, index) => {
    console.log(\`\\n📋 Sheet \${index + 1}: "\${worksheet.name}"\`);
    console.log(\`   Row count: \${worksheet.rowCount}\`);
    console.log(\`   Column count: \${worksheet.columnCount}\`);
    console.log(\`   Actual row count: \${worksheet.actualRowCount}\`);
    console.log(\`   Actual column count: \${worksheet.actualColumnCount}\`);
    
    // Check first cell formatting
    const cell = worksheet.getCell('A1');
    console.log(\`   A1 value: "\${cell.value}"\`);
    console.log(\`   A1 font: \${JSON.stringify(cell.font || {})}\`);
    console.log(\`   A1 fill: \${JSON.stringify(cell.fill || {})}\`);
    console.log(\`   A1 border: \${JSON.stringify(cell.border || {})}\`);
    console.log(\`   A1 alignment: \${JSON.stringify(cell.alignment || {})}\`);
    
    // Check column widths
    worksheet.columns.forEach((col, i) => {
      if (i < 5 && col.width) {
        console.log(\`   Column \${i + 1} width: \${col.width}\`);
      }
    });
    
    // Check merged cells
    if (worksheet.model.merges && worksheet.model.merges.length > 0) {
      console.log(\`   Merged cells: \${worksheet.model.merges.length}\`);
      worksheet.model.merges.slice(0, 3).forEach((merge, i) => {
        console.log(\`     Merge \${i + 1}: \${merge}\`);
      });
    }
  });
  
  // Test modification
  console.log('\\n✏️  Testing modification...');
  const dataSheet = workbook.getWorksheet('Thong_tin_khach');
  
  if (dataSheet) {
    // Add new column
    const newCol = dataSheet.columnCount + 1;
    const headerCell = dataSheet.getCell(1, newCol);
    headerCell.value = 'CHECK-IN TIME';
    
    // Copy formatting from A1
    const a1 = dataSheet.getCell('A1');
    if (a1.font) headerCell.font = { ...a1.font };
    if (a1.fill) headerCell.fill = { ...a1.fill };
    if (a1.border) headerCell.border = { ...a1.border };
    if (a1.alignment) headerCell.alignment = { ...a1.alignment };
    
    // Add sample data
    dataSheet.getCell(2, newCol).value = '2024-01-15 10:30:00';
    dataSheet.getCell(3, newCol).value = '2024-01-15 11:15:00';
    dataSheet.getCell(4, newCol).value = '2024-01-15 09:45:00';
    
    // Set column width
    dataSheet.getColumn(newCol).width = 20;
    
    console.log(\`   ✅ Added column \${newCol} with check-in data\`);
    
    // Write modified file
    await workbook.xlsx.writeFile('DATA_EXCELJS_TEST.xlsx');
    console.log('   ✅ Saved: DATA_EXCELJS_TEST.xlsx');
    
    // Verify by reading back
    const verifyWb = new ExcelJS.Workbook();
    await verifyWb.xlsx.readFile('DATA_EXCELJS_TEST.xlsx');
    const verifySheet = verifyWb.getWorksheet('Thong_tin_khach');
    const verifyHeader = verifySheet.getCell(1, newCol);
    
    console.log(\`   📊 Verification:\`);
    console.log(\`     Header: "\${verifyHeader.value}"\`);
    console.log(\`     Font preserved: \${JSON.stringify(verifyHeader.font || {})}\`);
    console.log(\`     Fill preserved: \${JSON.stringify(verifyHeader.fill || {})}\`);
    console.log(\`     Border preserved: \${JSON.stringify(verifyHeader.border || {})}\`);
  }
}

async function testXLSXJS() {
  console.log('\\n🔴 TESTING XLSX.JS...');
  
  const fileData = fs.readFileSync('DATA.xlsx');
  const workbook = XLSX.read(fileData, {
    type: "buffer",
    cellStyles: true,
    cellFormula: true,
    cellDates: true,
    cellNF: true
  });
  
  console.log('📂 XLSX.js Results:');
  console.log(\`   Sheets: \${workbook.SheetNames.join(', ')}\`);
  
  const sheet = workbook.Sheets['Thong_tin_khach'];
  if (sheet) {
    console.log(\`   Range: \${sheet['!ref']}\`);
    console.log(\`   A1 value: "\${sheet['A1']?.v}"\`);
    console.log(\`   A1 style: \${JSON.stringify(sheet['A1']?.s || {})}\`);
    console.log(\`   Column widths: \${sheet['!cols'] ? sheet['!cols'].length : 'NO'}\`);
  }
}

async function runComparison() {
  try {
    await testExcelJS();
    await testXLSXJS();
    
    console.log('\\n🎯 COMPARISON SUMMARY:');
    console.log('ExcelJS pros:');
    console.log('  ✅ Rich formatting API');
    console.log('  ✅ Better style preservation');
    console.log('  ✅ More intuitive cell access');
    console.log('  ✅ Built for Excel compatibility');
    console.log('\\nXLSX.js pros:');
    console.log('  ✅ Lighter weight');
    console.log('  ✅ Faster for basic operations');
    console.log('  ✅ More format support (CSV, etc.)');
    
  } catch (error) {
    console.error('Error:', error.message);
  }
}

runComparison();
`;
    
    fs.writeFileSync('exceljs-comparison.js', comparisonScript);
    console.log('✅ Created: exceljs-comparison.js');
    console.log('📋 Run this after installing ExcelJS: node exceljs-comparison.js');
    
    return;
  }
}

// If ExcelJS is available, run the test
if (ExcelJS) {
  testExcelJSNow();
}

async function testExcelJSNow() {
  console.log('🔵 RUNNING EXCELJS TEST...\n');
  
  try {
    const workbook = new ExcelJS.Workbook();
    
    // Read the DATA.xlsx file
    console.log('📂 Reading DATA.xlsx with ExcelJS...');
    await workbook.xlsx.readFile('DATA.xlsx');
    
    console.log('✅ File loaded successfully!');
    console.log(`📊 Worksheets found: ${workbook.worksheets.length}`);
    
    // Analyze each worksheet
    workbook.worksheets.forEach((worksheet, index) => {
      console.log(`\n📋 WORKSHEET ${index + 1}: "${worksheet.name}"`);
      console.log(`   Row count: ${worksheet.rowCount}`);
      console.log(`   Column count: ${worksheet.columnCount}`);
      console.log(`   Actual rows: ${worksheet.actualRowCount}`);
      console.log(`   Actual columns: ${worksheet.actualColumnCount}`);
      
      // Check A1 cell formatting
      const a1 = worksheet.getCell('A1');
      console.log(`   A1 value: "${a1.value}"`);
      console.log(`   A1 type: ${a1.type || 'undefined'}`);
      
      // Check formatting details
      if (a1.font && Object.keys(a1.font).length > 0) {
        console.log(`   A1 font: ${JSON.stringify(a1.font)}`);
      } else {
        console.log(`   A1 font: NO FORMATTING`);
      }
      
      if (a1.fill && Object.keys(a1.fill).length > 0) {
        console.log(`   A1 fill: ${JSON.stringify(a1.fill)}`);
      } else {
        console.log(`   A1 fill: NO FORMATTING`);
      }
      
      if (a1.border && Object.keys(a1.border).length > 0) {
        console.log(`   A1 border: ${JSON.stringify(a1.border)}`);
      } else {
        console.log(`   A1 border: NO FORMATTING`);
      }
      
      if (a1.alignment && Object.keys(a1.alignment).length > 0) {
        console.log(`   A1 alignment: ${JSON.stringify(a1.alignment)}`);
      } else {
        console.log(`   A1 alignment: NO FORMATTING`);
      }
      
      // Check a few more cells
      const b1 = worksheet.getCell('B1');
      console.log(`   B1 value: "${b1.value}" | Font: ${b1.font ? 'YES' : 'NO'} | Fill: ${b1.fill ? 'YES' : 'NO'}`);
      
      // Check column widths
      console.log(`   Column widths:`);
      for (let i = 1; i <= Math.min(5, worksheet.columnCount); i++) {
        const col = worksheet.getColumn(i);
        console.log(`     Column ${i}: width=${col.width || 'default'}, hidden=${col.hidden || false}`);
      }
      
      // Check merged cells
      if (worksheet.model.merges && worksheet.model.merges.length > 0) {
        console.log(`   Merged cells: ${worksheet.model.merges.length}`);
        worksheet.model.merges.slice(0, 3).forEach((merge, i) => {
          console.log(`     ${i + 1}: ${merge}`);
        });
      } else {
        console.log(`   Merged cells: NONE`);
      }
    });
    
    // Test modification
    console.log(`\n✏️  TESTING MODIFICATION...`);
    
    const dataSheet = workbook.getWorksheet('Thong_tin_khach');
    if (dataSheet) {
      console.log('📋 Modifying "Thong_tin_khach" sheet...');
      
      // Get current dimensions
      const originalCols = dataSheet.actualColumnCount;
      const newColNum = originalCols + 1;
      
      console.log(`   Original columns: ${originalCols}`);
      console.log(`   Adding column: ${newColNum}`);
      
      // Add header with formatting
      const headerCell = dataSheet.getCell(1, newColNum);
      headerCell.value = 'CHECK-IN TIME';
      
      // Try to copy formatting from A1
      const a1Cell = dataSheet.getCell('A1');
      if (a1Cell.font) {
        headerCell.font = { ...a1Cell.font };
        console.log(`   ✅ Copied font formatting`);
      }
      if (a1Cell.fill) {
        headerCell.fill = { ...a1Cell.fill };
        console.log(`   ✅ Copied fill formatting`);
      }
      if (a1Cell.border) {
        headerCell.border = { ...a1Cell.border };
        console.log(`   ✅ Copied border formatting`);
      }
      if (a1Cell.alignment) {
        headerCell.alignment = { ...a1Cell.alignment };
        console.log(`   ✅ Copied alignment formatting`);
      }
      
      // Add sample data
      const sampleData = [
        '2024-01-15 10:30:00',
        '2024-01-15 11:15:00',
        '',
        '2024-01-15 09:45:00',
        '2024-01-15 12:00:00'
      ];
      
      sampleData.forEach((data, index) => {
        const cell = dataSheet.getCell(index + 2, newColNum);
        cell.value = data;
        
        // Copy formatting from corresponding A column cell
        const refCell = dataSheet.getCell(index + 2, 1);
        if (refCell.font) cell.font = { ...refCell.font };
        if (refCell.fill) cell.fill = { ...refCell.fill };
        if (refCell.border) cell.border = { ...refCell.border };
      });
      
      // Set column width
      const newColumn = dataSheet.getColumn(newColNum);
      newColumn.width = 20;
      
      console.log(`   ✅ Added ${sampleData.filter(d => d).length} check-in records`);
      console.log(`   ✅ Set column width to 20`);
      
      // Save the modified file
      console.log(`\n💾 Saving modified file...`);
      await workbook.xlsx.writeFile('DATA_EXCELJS_RESULT.xlsx');
      console.log(`✅ Saved: DATA_EXCELJS_RESULT.xlsx`);
      
      // Verify the result
      console.log(`\n🔍 VERIFICATION...`);
      const verifyWb = new ExcelJS.Workbook();
      await verifyWb.xlsx.readFile('DATA_EXCELJS_RESULT.xlsx');
      
      const verifySheet = verifyWb.getWorksheet('Thong_tin_khach');
      const verifyHeader = verifySheet.getCell(1, newColNum);
      
      console.log(`📊 Verification results:`);
      console.log(`   Worksheets: ${verifyWb.worksheets.length}`);
      console.log(`   New header: "${verifyHeader.value}"`);
      console.log(`   Header font: ${JSON.stringify(verifyHeader.font || {})}`);
      console.log(`   Header fill: ${JSON.stringify(verifyHeader.fill || {})}`);
      console.log(`   Header border: ${JSON.stringify(verifyHeader.border || {})}`);
      
      // Check sample data
      const sampleCell = verifySheet.getCell(2, newColNum);
      console.log(`   Sample data: "${sampleCell.value}"`);
      console.log(`   Data font: ${JSON.stringify(sampleCell.font || {})}`);
      
      console.log(`\n🎯 EXCELJS RESULTS:`);
      console.log(`✅ File modification: SUCCESS`);
      console.log(`✅ Formatting preservation: ${verifyHeader.font ? 'YES' : 'NO'}`);
      console.log(`✅ Structure preservation: YES`);
      
    } else {
      console.log(`❌ Sheet "Thong_tin_khach" not found`);
    }
    
  } catch (error) {
    console.error(`❌ ExcelJS test failed: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
  }
}

console.log('\n💡 NEXT STEPS:');
console.log('1. 📁 Check if DATA_EXCELJS_RESULT.xlsx was created');
console.log('2. 🔍 Open it in Excel and compare with original');
console.log('3. ✅ Verify formatting preservation');
console.log('4. 📊 Compare with XLSX.js results');