
// ExcelJS vs XLSX.js Comparison Script
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const fs = require('fs');

console.log('📊 EXCELJS VS XLSX.JS COMPARISON\n');

async function testExcelJS() {
  console.log('🔵 TESTING EXCELJS...');
  
  const workbook = new ExcelJS.Workbook();
  
  // Read the file
  await workbook.xlsx.readFile('DATA.xlsx');
  
  console.log('📂 ExcelJS Results:');
  console.log(`   Worksheets: ${workbook.worksheets.length}`);
  
  workbook.worksheets.forEach((worksheet, index) => {
    console.log(`\n📋 Sheet ${index + 1}: "${worksheet.name}"`);
    console.log(`   Row count: ${worksheet.rowCount}`);
    console.log(`   Column count: ${worksheet.columnCount}`);
    console.log(`   Actual row count: ${worksheet.actualRowCount}`);
    console.log(`   Actual column count: ${worksheet.actualColumnCount}`);
    
    // Check first cell formatting
    const cell = worksheet.getCell('A1');
    console.log(`   A1 value: "${cell.value}"`);
    console.log(`   A1 font: ${JSON.stringify(cell.font || {})}`);
    console.log(`   A1 fill: ${JSON.stringify(cell.fill || {})}`);
    console.log(`   A1 border: ${JSON.stringify(cell.border || {})}`);
    console.log(`   A1 alignment: ${JSON.stringify(cell.alignment || {})}`);
    
    // Check column widths
    worksheet.columns.forEach((col, i) => {
      if (i < 5 && col.width) {
        console.log(`   Column ${i + 1} width: ${col.width}`);
      }
    });
    
    // Check merged cells
    if (worksheet.model.merges && worksheet.model.merges.length > 0) {
      console.log(`   Merged cells: ${worksheet.model.merges.length}`);
      worksheet.model.merges.slice(0, 3).forEach((merge, i) => {
        console.log(`     Merge ${i + 1}: ${merge}`);
      });
    }
  });
  
  // Test modification
  console.log('\n✏️  Testing modification...');
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
    
    console.log(`   ✅ Added column ${newCol} with check-in data`);
    
    // Write modified file
    await workbook.xlsx.writeFile('DATA_EXCELJS_TEST.xlsx');
    console.log('   ✅ Saved: DATA_EXCELJS_TEST.xlsx');
    
    // Verify by reading back
    const verifyWb = new ExcelJS.Workbook();
    await verifyWb.xlsx.readFile('DATA_EXCELJS_TEST.xlsx');
    const verifySheet = verifyWb.getWorksheet('Thong_tin_khach');
    const verifyHeader = verifySheet.getCell(1, newCol);
    
    console.log(`   📊 Verification:`);
    console.log(`     Header: "${verifyHeader.value}"`);
    console.log(`     Font preserved: ${JSON.stringify(verifyHeader.font || {})}`);
    console.log(`     Fill preserved: ${JSON.stringify(verifyHeader.fill || {})}`);
    console.log(`     Border preserved: ${JSON.stringify(verifyHeader.border || {})}`);
  }
}

async function testXLSXJS() {
  console.log('\n🔴 TESTING XLSX.JS...');
  
  const fileData = fs.readFileSync('DATA.xlsx');
  const workbook = XLSX.read(fileData, {
    type: "buffer",
    cellStyles: true,
    cellFormula: true,
    cellDates: true,
    cellNF: true
  });
  
  console.log('📂 XLSX.js Results:');
  console.log(`   Sheets: ${workbook.SheetNames.join(', ')}`);
  
  const sheet = workbook.Sheets['Thong_tin_khach'];
  if (sheet) {
    console.log(`   Range: ${sheet['!ref']}`);
    console.log(`   A1 value: "${sheet['A1']?.v}"`);
    console.log(`   A1 style: ${JSON.stringify(sheet['A1']?.s || {})}`);
    console.log(`   Column widths: ${sheet['!cols'] ? sheet['!cols'].length : 'NO'}`);
  }
}

async function runComparison() {
  try {
    await testExcelJS();
    await testXLSXJS();
    
    console.log('\n🎯 COMPARISON SUMMARY:');
    console.log('ExcelJS pros:');
    console.log('  ✅ Rich formatting API');
    console.log('  ✅ Better style preservation');
    console.log('  ✅ More intuitive cell access');
    console.log('  ✅ Built for Excel compatibility');
    console.log('\nXLSX.js pros:');
    console.log('  ✅ Lighter weight');
    console.log('  ✅ Faster for basic operations');
    console.log('  ✅ More format support (CSV, etc.)');
    
  } catch (error) {
    console.error('Error:', error.message);
  }
}

runComparison();
