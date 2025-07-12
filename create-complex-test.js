// Create a complex Excel file with extensive formatting to test our utilities
const XLSX = require('xlsx');
const fs = require('fs');

// Inline utility functions for testing
const deepCopyWorkbook = (workbook) => {
  try {
    const newWorkbook = XLSX.utils.book_new();
    
    Object.keys(workbook.Sheets).forEach(sheetName => {
      try {
        const originalSheet = workbook.Sheets[sheetName];
        const newSheet = {};
        
        Object.keys(originalSheet).forEach(key => {
          try {
            if (key.startsWith('!')) {
              const value = originalSheet[key];
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
              const cell = originalSheet[key];
              if (cell && typeof cell === 'object') {
                newSheet[key] = JSON.parse(JSON.stringify(cell));
              }
            }
          } catch (error) {
            newSheet[key] = originalSheet[key];
          }
        });
        
        newWorkbook.Sheets[sheetName] = newSheet;
      } catch (error) {
        newWorkbook.Sheets[sheetName] = workbook.Sheets[sheetName];
      }
    });
    
    newWorkbook.SheetNames = [...workbook.SheetNames];
    if (workbook.Props) newWorkbook.Props = JSON.parse(JSON.stringify(workbook.Props));
    if (workbook.Custprops) newWorkbook.Custprops = JSON.parse(JSON.stringify(workbook.Custprops));
    if (workbook.Workbook) newWorkbook.Workbook = JSON.parse(JSON.stringify(workbook.Workbook));
    if (workbook.vbaraw) newWorkbook.vbaraw = workbook.vbaraw;
    
    return newWorkbook;
  } catch (error) {
    return workbook;
  }
};

const detectFileFormatting = (workbook) => {
  try {
    if (!workbook || !workbook.SheetNames) return false;

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet || typeof sheet !== 'object') continue;
      
      if (sheet['!cols'] || sheet['!rows'] || sheet['!merges'] || 
          sheet['!protect'] || sheet['!autofilter'] || sheet['!margins']) {
        return true;
      }
      
      for (const cellAddress of Object.keys(sheet)) {
        if (!cellAddress.startsWith('!')) {
          const cell = sheet[cellAddress];
          if (cell && typeof cell === 'object' && 
              cell !== null && !Array.isArray(cell) &&
              (cell.hasOwnProperty('s') || cell.hasOwnProperty('z') || 
               cell.hasOwnProperty('l') || cell.hasOwnProperty('c'))) {
            return true;
          }
        }
      }
    }
  } catch (error) {
    return true;
  }
  return false;
};

function createComplexFormattedExcel() {
  const wb = XLSX.utils.book_new();
  
  // Create complex data
  const data = [
    ['🎫 ATTENDEE REGISTRATION SYSTEM', '', '', '', '', ''],
    ['Event: Tech Conference 2024', '', 'Date: Jan 15-16, 2024', '', '', ''],
    ['', '', '', '', '', ''],
    ['ID', 'Full Name', 'Email Address', 'Phone Number', 'Department', 'VIP Status'],
    ['001', 'John Smith', 'john.smith@company.com', '+1-555-0101', 'Engineering', 'Yes'],
    ['002', 'Sarah Johnson', 'sarah.j@company.com', '+1-555-0102', 'Marketing', 'No'],
    ['003', 'Michael Brown', 'michael.brown@company.com', '+1-555-0103', 'Sales', 'Yes'],
    ['004', 'Emily Davis', 'emily.davis@company.com', '+1-555-0104', 'HR', 'No'],
    ['005', 'David Wilson', 'david.wilson@company.com', '+1-555-0105', 'Finance', 'Yes'],
    ['', '', '', '', '', ''],
    ['SUMMARY:', '', '', '', '', ''],
    ['Total Attendees:', '5', 'VIP Count:', '3', 'Regular:', '2']
  ];
  
  const ws = XLSX.utils.aoa_to_sheet(data);
  
  // Set comprehensive formatting
  
  // Column widths
  ws['!cols'] = [
    { width: 8 },   // ID
    { width: 20 },  // Name
    { width: 25 },  // Email
    { width: 18 },  // Phone
    { width: 15 },  // Department
    { width: 12 }   // VIP Status
  ];
  
  // Row heights
  ws['!rows'] = [
    { hpt: 25 }, // Title row
    { hpt: 20 }, // Subtitle row
    { hpt: 15 }, // Empty row
    { hpt: 20 }, // Header row
    null, null, null, null, null, // Data rows
    { hpt: 15 }, // Empty row
    { hpt: 20 }, // Summary header
    { hpt: 18 }  // Summary data
  ];
  
  // Merged cells
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }, // Title
    { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }, // Event info
    { s: { r: 1, c: 2 }, e: { r: 1, c: 5 } }, // Date info
    { s: { r: 10, c: 0 }, e: { r: 10, c: 1 } }, // Summary label
  ];
  
  // Cell styling
  
  // Title row (A1)
  ws['A1'].s = {
    font: { 
      bold: true, 
      sz: 16, 
      color: { rgb: 'FFFFFF' } 
    },
    fill: { 
      fgColor: { rgb: '2E75B6' } 
    },
    alignment: { 
      horizontal: 'center', 
      vertical: 'center' 
    },
    border: {
      top: { style: 'thick', color: { rgb: '000000' } },
      bottom: { style: 'thick', color: { rgb: '000000' } },
      left: { style: 'thick', color: { rgb: '000000' } },
      right: { style: 'thick', color: { rgb: '000000' } }
    }
  };
  
  // Event info (A2)
  ws['A2'].s = {
    font: { bold: true, sz: 12 },
    fill: { fgColor: { rgb: 'D9E1F2' } },
    alignment: { horizontal: 'left' }
  };
  
  // Date info (C2)
  ws['C2'].s = {
    font: { bold: true, sz: 12 },
    fill: { fgColor: { rgb: 'D9E1F2' } },
    alignment: { horizontal: 'center' }
  };
  
  // Header row (row 4)
  ['A4', 'B4', 'C4', 'D4', 'E4', 'F4'].forEach(cell => {
    if (ws[cell]) {
      ws[cell].s = {
        font: { 
          bold: true, 
          color: { rgb: 'FFFFFF' },
          sz: 11
        },
        fill: { 
          fgColor: { rgb: '4472C4' } 
        },
        alignment: { 
          horizontal: 'center',
          vertical: 'center'
        },
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        }
      };
    }
  });
  
  // Data rows with alternating colors
  for (let row = 5; row <= 9; row++) {
    const isEvenRow = row % 2 === 0;
    const fillColor = isEvenRow ? 'F2F2F2' : 'FFFFFF';
    
    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
      const cellAddr = `${col}${row}`;
      if (ws[cellAddr]) {
        ws[cellAddr].s = {
          fill: { fgColor: { rgb: fillColor } },
          border: {
            top: { style: 'thin', color: { rgb: 'CCCCCC' } },
            bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
            left: { style: 'thin', color: { rgb: 'CCCCCC' } },
            right: { style: 'thin', color: { rgb: 'CCCCCC' } }
          },
          alignment: { 
            horizontal: col === 'A' ? 'center' : 'left',
            vertical: 'center'
          }
        };
        
        // Special formatting for VIP status
        if (col === 'F' && ws[cellAddr].v === 'Yes') {
          ws[cellAddr].s.font = { 
            bold: true, 
            color: { rgb: '00AA00' } 
          };
        } else if (col === 'F' && ws[cellAddr].v === 'No') {
          ws[cellAddr].s.font = { 
            color: { rgb: 'AA0000' } 
          };
        }
      }
    });
  }
  
  // Summary section formatting
  if (ws['A11']) {
    ws['A11'].s = {
      font: { bold: true, sz: 12 },
      fill: { fgColor: { rgb: 'FFE699' } }
    };
  }
  
  ['A12', 'B12', 'C12', 'D12', 'E12', 'F12'].forEach(cell => {
    if (ws[cell]) {
      ws[cell].s = {
        font: { bold: true },
        fill: { fgColor: { rgb: 'FFF2CC' } },
        alignment: { horizontal: 'center' }
      };
    }
  });
  
  // Number formatting for phone numbers
  for (let row = 5; row <= 9; row++) {
    const phoneCell = `D${row}`;
    if (ws[phoneCell]) {
      ws[phoneCell].z = '[>=0]"+"0-000-0000;[<0]"+"0-000-0000'; // Custom phone format
    }
  }
  
  XLSX.utils.book_append_sheet(wb, ws, 'Attendees');
  
  // Add a second sheet with different formatting
  const summaryData = [
    ['DEPARTMENT BREAKDOWN', ''],
    ['Department', 'Count'],
    ['Engineering', '1'],
    ['Marketing', '1'],
    ['Sales', '1'],
    ['HR', '1'],
    ['Finance', '1']
  ];
  
  const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
  
  // Format summary sheet
  summaryWs['!cols'] = [{ width: 15 }, { width: 10 }];
  summaryWs['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
  
  // Header formatting
  summaryWs['A1'].s = {
    font: { bold: true, sz: 14, color: { rgb: 'FFFFFF' } },
    fill: { fgColor: { rgb: 'E74C3C' } },
    alignment: { horizontal: 'center' }
  };
  
  ['A2', 'B2'].forEach(cell => {
    if (summaryWs[cell]) {
      summaryWs[cell].s = {
        font: { bold: true },
        fill: { fgColor: { rgb: 'FADBD8' } },
        alignment: { horizontal: 'center' }
      };
    }
  });
  
  XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');
  
  return wb;
}

// Test the complex file
console.log('🏗️  Creating complex formatted Excel file...');

const complexWb = createComplexFormattedExcel();

// Save original file
const originalBuffer = XLSX.write(complexWb, { 
  type: 'array', 
  bookType: 'xlsx',
  cellStyles: true 
});

fs.writeFileSync('complex-original.xlsx', Buffer.from(originalBuffer));
console.log('✅ Created complex-original.xlsx');

// Test our utilities with the complex file
console.log('\n🧪 Testing utilities with complex file...');

// Test detection
const hasFormatting = detectFileFormatting(complexWb);
console.log(`📊 Formatting detected: ${hasFormatting ? '✅ YES' : '❌ NO'}`);

// Test deep copy
const copiedWb = deepCopyWorkbook(complexWb);
console.log('📋 Deep copy created');

// Simulate adding check-in data
const attendeesSheet = copiedWb.Sheets['Attendees'];
const currentRange = XLSX.utils.decode_range(attendeesSheet['!ref'] || 'A1:F12');

// Add check-in column
const newColIndex = currentRange.e.c + 1;
const newColLetter = XLSX.utils.encode_col(newColIndex);

// Header
attendeesSheet[`${newColLetter}4`] = {
  v: 'Check-in Time',
  t: 's',
  s: {
    font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
    fill: { fgColor: { rgb: '4472C4' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } }
    }
  }
};

// Check-in times (simulate some people checked in)
const checkInTimes = ['10:30 AM', '', '11:15 AM', '', '09:45 AM'];

for (let i = 0; i < checkInTimes.length; i++) {
  const row = i + 5;
  const cellAddr = `${newColLetter}${row}`;
  const isEvenRow = row % 2 === 0;
  const fillColor = isEvenRow ? 'F2F2F2' : 'FFFFFF';
  
  attendeesSheet[cellAddr] = {
    v: checkInTimes[i],
    t: 's',
    s: {
      fill: { fgColor: { rgb: fillColor } },
      border: {
        top: { style: 'thin', color: { rgb: 'CCCCCC' } },
        bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
        left: { style: 'thin', color: { rgb: 'CCCCCC' } },
        right: { style: 'thin', color: { rgb: 'CCCCCC' } }
      },
      alignment: { horizontal: 'center', vertical: 'center' },
      font: checkInTimes[i] ? { color: { rgb: '00AA00' }, bold: true } : undefined
    }
  };
}

// Update range
attendeesSheet['!ref'] = XLSX.utils.encode_range({
  s: { r: 0, c: 0 },
  e: { r: currentRange.e.r, c: newColIndex }
});

// Update column widths
if (attendeesSheet['!cols']) {
  attendeesSheet['!cols'].push({ width: 15 });
} else {
  attendeesSheet['!cols'] = [...Array(newColIndex).fill({ width: 15 }), { width: 15 }];
}

// Write modified file
const modifiedBuffer = XLSX.write(copiedWb, {
  type: 'array',
  bookType: 'xlsx',
  cellStyles: true
});

fs.writeFileSync('complex-modified.xlsx', Buffer.from(modifiedBuffer));
console.log('✅ Created complex-modified.xlsx with check-in data');

// Verify the modified file still has formatting
const rereadWb = XLSX.read(modifiedBuffer, {
  type: 'array',
  cellStyles: true,
  cellFormula: true,
  cellDates: true,
  cellNF: true
});

const stillHasFormatting = detectFileFormatting(rereadWb);
console.log(`📊 Formatting preserved after modification: ${stillHasFormatting ? '✅ YES' : '❌ NO'}`);

// Check specific formatting elements
const modifiedSheet = rereadWb.Sheets['Attendees'];
const hasColumns = !!modifiedSheet['!cols'];
const hasMerges = !!modifiedSheet['!merges'];
const hasColoredHeader = !!modifiedSheet['A1']?.s?.fill;
const hasNewColumn = !!modifiedSheet['G4'];

console.log('\n📋 Detailed verification:');
console.log(`   Column widths preserved: ${hasColumns ? '✅' : '❌'}`);
console.log(`   Merged cells preserved: ${hasMerges ? '✅' : '❌'}`);
console.log(`   Header colors preserved: ${hasColoredHeader ? '✅' : '❌'}`);
console.log(`   New check-in column added: ${hasNewColumn ? '✅' : '❌'}`);

if (hasColumns && hasMerges && hasColoredHeader && hasNewColumn) {
  console.log('\n🎉 SUCCESS! All formatting preserved and new data added correctly!');
  console.log('\n📁 Files created:');
  console.log('   • complex-original.xlsx - Original file with extensive formatting');
  console.log('   • complex-modified.xlsx - Modified file with check-in data and preserved formatting');
  console.log('\n💡 Open both files in Excel to visually verify formatting preservation');
} else {
  console.log('\n⚠️  Some formatting may have been lost. Check the files manually.');
}