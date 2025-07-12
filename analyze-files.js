// Analyze the created Excel files to understand what formatting is preserved
const XLSX = require('xlsx');
const fs = require('fs');

function analyzeExcelFile(filename) {
  console.log(`\n📊 Analyzing ${filename}:`);
  
  if (!fs.existsSync(filename)) {
    console.log(`❌ File ${filename} does not exist`);
    return;
  }
  
  const buffer = fs.readFileSync(filename);
  const wb = XLSX.read(buffer, {
    type: 'buffer',
    cellStyles: true,
    cellFormula: true,
    cellDates: true,
    cellNF: true
  });
  
  console.log(`   Sheets: ${wb.SheetNames.join(', ')}`);
  
  const sheet = wb.Sheets[wb.SheetNames[0]];
  
  // Check sheet-level formatting
  console.log(`   Column widths: ${sheet['!cols'] ? '✅ Present' : '❌ Missing'}`);
  console.log(`   Row heights: ${sheet['!rows'] ? '✅ Present' : '❌ Missing'}`);
  console.log(`   Merged cells: ${sheet['!merges'] ? '✅ Present' : '❌ Missing'}`);
  console.log(`   Range: ${sheet['!ref'] || 'N/A'}`);
  
  // Check specific cell formatting
  const cellsToCheck = ['A1', 'A4', 'B4', 'C4', 'G4'];
  cellsToCheck.forEach(cellAddr => {
    const cell = sheet[cellAddr];
    if (cell) {
      console.log(`   ${cellAddr}: value="${cell.v}", has style=${!!cell.s}`);
      if (cell.s) {
        console.log(`     - Font: ${JSON.stringify(cell.s.font || {})}`);
        console.log(`     - Fill: ${JSON.stringify(cell.s.fill || {})}`);
        console.log(`     - Alignment: ${JSON.stringify(cell.s.alignment || {})}`);
      }
    } else {
      console.log(`   ${cellAddr}: Not found`);
    }
  });
  
  // Count cells with formatting
  let cellsWithStyles = 0;
  let totalCells = 0;
  
  Object.keys(sheet).forEach(key => {
    if (!key.startsWith('!')) {
      totalCells++;
      if (sheet[key].s) {
        cellsWithStyles++;
      }
    }
  });
  
  console.log(`   Cells with styles: ${cellsWithStyles}/${totalCells}`);
}

// Analyze both files
analyzeExcelFile('complex-original.xlsx');
analyzeExcelFile('complex-modified.xlsx');

// Create a simple test to understand what's happening
console.log('\n🔬 Simple format preservation test:');

const simpleWb = XLSX.utils.book_new();
const simpleWs = XLSX.utils.aoa_to_sheet([
  ['Styled Header', 'Normal'],
  ['Data1', 'Data2']
]);

// Add simple styling
simpleWs['A1'].s = {
  font: { bold: true, color: { rgb: 'FFFFFF' } },
  fill: { fgColor: { rgb: 'FF0000' } }
};

simpleWs['!cols'] = [{ width: 15 }, { width: 10 }];

XLSX.utils.book_append_sheet(simpleWb, simpleWs, 'Simple');

// Write and read back
const simpleBuffer = XLSX.write(simpleWb, {
  type: 'array',
  bookType: 'xlsx',
  cellStyles: true
});

fs.writeFileSync('simple-test.xlsx', Buffer.from(simpleBuffer));

const rereadWb = XLSX.read(simpleBuffer, {
  type: 'array',
  cellStyles: true,
  cellFormula: true,
  cellDates: true,
  cellNF: true
});

const rereadWs = rereadWb.Sheets['Simple'];

console.log('Original A1 style:', JSON.stringify(simpleWs['A1'].s, null, 2));
console.log('Re-read A1 style:', JSON.stringify(rereadWs['A1']?.s || 'NO STYLE', null, 2));
console.log('Columns preserved:', !!rereadWs['!cols']);

analyzeExcelFile('simple-test.xlsx');