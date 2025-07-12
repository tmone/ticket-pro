/**
 * @jest-environment jsdom
 */
import * as XLSX from 'xlsx';
import { format } from 'date-fns';
import {
  deepCopyWorkbook,
  detectFileFormatting,
  readExcelWithFormatting,
  writeExcelWithFormatting,
  addColumnWithFormatting,
  deepCloneWorksheet
} from '../src/lib/excel-utils';

// Create a realistic test Excel file with formatting
const createRealisticExcelFile = (): Uint8Array => {
  const wb = XLSX.utils.book_new();
  
  // Create worksheet with real data
  const data = [
    ['Name', 'Email', 'Phone', 'Registration Date'],
    ['John Doe', 'john@example.com', '123-456-7890', '2024-01-15'],
    ['Jane Smith', 'jane@example.com', '098-765-4321', '2024-01-16'],
    ['Bob Johnson', 'bob@example.com', '555-123-4567', '2024-01-17']
  ];
  
  const ws = XLSX.utils.aoa_to_sheet(data);
  
  // Add comprehensive formatting
  ws['!cols'] = [
    { width: 20 },  // Name
    { width: 25 },  // Email
    { width: 18 },  // Phone
    { width: 18 }   // Registration Date
  ];
  
  // Add merged cells
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } } // Merge header row
  ];
  
  // Add cell formatting
  if (ws['A1']) {
    ws['A1'] = {
      ...ws['A1'],
      s: {
        font: { bold: true, sz: 14, color: { rgb: 'FFFFFF' } },
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
  }
  
  // Format header cells
  ['A1', 'B1', 'C1', 'D1'].forEach(cell => {
    if (ws[cell]) {
      ws[cell].s = {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '4472C4' } },
        alignment: { horizontal: 'center' }
      };
    }
  });
  
  // Format data cells with alternating row colors
  for (let row = 2; row <= 4; row++) {
    const isEvenRow = row % 2 === 0;
    const fillColor = isEvenRow ? 'F2F2F2' : 'FFFFFF';
    
    ['A', 'B', 'C', 'D'].forEach(col => {
      const cellAddr = `${col}${row}`;
      if (ws[cellAddr]) {
        ws[cellAddr].s = {
          fill: { fgColor: { rgb: fillColor } },
          border: {
            top: { style: 'thin', color: { rgb: 'CCCCCC' } },
            bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
            left: { style: 'thin', color: { rgb: 'CCCCCC' } },
            right: { style: 'thin', color: { rgb: 'CCCCCC' } }
          }
        };
      }
    });
  }
  
  // Add data validation and number formats
  if (ws['D2']) {
    ws['D2'].z = 'yyyy-mm-dd'; // Date format
  }
  
  XLSX.utils.book_append_sheet(wb, ws, 'Attendees');
  
  return XLSX.write(wb, { 
    bookType: 'xlsx', 
    type: 'array',
    cellStyles: true 
  });
};

describe('Excel Integration Tests', () => {
  let originalFileData: Uint8Array;
  let workbook: XLSX.WorkBook;

  beforeEach(() => {
    originalFileData = createRealisticExcelFile();
    workbook = readExcelWithFormatting(originalFileData);
  });

  describe('Full File Processing Workflow', () => {
    it('should preserve all formatting through upload-modify-download cycle', () => {
      // Step 1: Verify original file has formatting
      expect(detectFileFormatting(workbook)).toBe(true);
      
      const originalSheet = workbook.Sheets['Attendees'];
      expect(originalSheet['!cols']).toBeDefined();
      expect(originalSheet['!merges']).toBeDefined();
      expect(originalSheet['A1']?.s).toBeDefined();
      
      // Step 2: Process data (simulate user interaction)
      const jsonData = XLSX.utils.sheet_to_json(originalSheet, { defval: '' });
      expect(jsonData).toHaveLength(3); // 3 data rows (excluding header)
      
      // Step 3: Simulate check-in data
      const processedRows = jsonData.map((row: any, index) => ({
        ...row,
        __rowNum__: index + 2,
        checkedInTime: index === 0 ? new Date('2024-01-15T10:30:00') : null
      }));
      
      // Step 4: Create modified workbook
      const modifiedWorkbook = deepCopyWorkbook(workbook);
      const clonedSheet = deepCloneWorksheet(originalSheet);
      
      // Step 5: Add check-in column
      const checkInValues = processedRows.map(row => 
        row.checkedInTime ? format(new Date(row.checkedInTime), 'yyyy-MM-dd HH:mm:ss') : ''
      );
      
      const updatedSheet = addColumnWithFormatting(clonedSheet, {
        header: 'Checked-In At',
        values: checkInValues
      });
      
      modifiedWorkbook.Sheets['Attendees'] = updatedSheet;
      
      // Step 6: Export to buffer
      const exportedBuffer = writeExcelWithFormatting(modifiedWorkbook);
      
      // Step 7: Re-read exported file and verify
      const rereadWorkbook = readExcelWithFormatting(exportedBuffer);
      const rereadSheet = rereadWorkbook.Sheets['Attendees'];
      
      // Verify formatting is preserved
      expect(detectFileFormatting(rereadWorkbook)).toBe(true);
      expect(rereadSheet['!cols']).toBeDefined();
      expect(rereadSheet['!cols']).toHaveLength(5); // Original 4 + new 1
      expect(rereadSheet['!merges']).toBeDefined();
      
      // Verify original cell formatting is preserved
      expect(rereadSheet['A1']?.s).toBeDefined();
      expect(rereadSheet['A1']?.s?.font?.bold).toBe(true);
      expect(rereadSheet['A1']?.s?.fill?.fgColor?.rgb).toBe('4472C4');
      
      // Verify new column data
      expect(rereadSheet['E1']?.v).toBe('Checked-In At');
      expect(rereadSheet['E2']?.v).toBe('2024-01-15 10:30:00');
      expect(rereadSheet['E3']?.v).toBe('');
      expect(rereadSheet['E4']?.v).toBe('');
      
      // Verify range is updated correctly
      const finalRange = XLSX.utils.decode_range(rereadSheet['!ref'] || '');
      expect(finalRange.e.c).toBe(4); // 0-indexed, so column E is 4
    });

    it('should handle multiple sheets with different formatting', () => {
      // Create workbook with multiple sheets
      const multiSheetWb = XLSX.utils.book_new();
      
      // Sheet 1: Attendees (formatted)
      const sheet1 = XLSX.utils.aoa_to_sheet([
        ['Name', 'Status'],
        ['John', 'Active'],
        ['Jane', 'Inactive']
      ]);
      sheet1['!cols'] = [{ width: 15 }, { width: 10 }];
      sheet1['A1'].s = { font: { bold: true } };
      
      // Sheet 2: Settings (different formatting)
      const sheet2 = XLSX.utils.aoa_to_sheet([
        ['Setting', 'Value'],
        ['Theme', 'Dark'],
        ['Language', 'EN']
      ]);
      sheet2['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
      sheet2['A1'].s = { fill: { fgColor: { rgb: 'FF0000' } } };
      
      XLSX.utils.book_append_sheet(multiSheetWb, sheet1, 'Attendees');
      XLSX.utils.book_append_sheet(multiSheetWb, sheet2, 'Settings');
      
      // Process through our utilities
      const copiedWb = deepCopyWorkbook(multiSheetWb);
      
      // Modify only one sheet
      const modifiedSheet = addColumnWithFormatting(copiedWb.Sheets['Attendees'], {
        header: 'Check-in',
        values: ['10:30', '']
      });
      copiedWb.Sheets['Attendees'] = modifiedSheet;
      
      // Export and re-read
      const buffer = writeExcelWithFormatting(copiedWb);
      const finalWb = readExcelWithFormatting(buffer);
      
      // Verify both sheets preserved their formatting
      expect(finalWb.SheetNames).toEqual(['Attendees', 'Settings']);
      expect(finalWb.Sheets['Attendees']['!cols']).toBeDefined();
      expect(finalWb.Sheets['Settings']['!merges']).toBeDefined();
      expect(finalWb.Sheets['Attendees']['A1']?.s?.font?.bold).toBe(true);
      expect(finalWb.Sheets['Settings']['A1']?.s?.fill?.fgColor?.rgb).toBe('FF0000');
      
      // Verify new column only in modified sheet
      expect(finalWb.Sheets['Attendees']['C1']?.v).toBe('Check-in');
      expect(finalWb.Sheets['Settings']['C1']).toBeUndefined();
    });

    it('should preserve VBA and custom properties', () => {
      // Create workbook with custom properties
      const wb = XLSX.utils.book_new();
      wb.Props = {
        Title: 'Attendee List',
        Author: 'Test System',
        CreatedDate: new Date('2024-01-01')
      };
      wb.Custprops = {
        Department: 'IT',
        Version: '1.0'
      };
      
      const ws = XLSX.utils.aoa_to_sheet([['Name'], ['John']]);
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      
      // Process through utilities
      const processed = deepCopyWorkbook(wb);
      
      // Verify properties are preserved
      expect(processed.Props?.Title).toBe('Attendee List');
      expect(processed.Props?.Author).toBe('Test System');
      expect(processed.Custprops?.Department).toBe('IT');
      expect(processed.Custprops?.Version).toBe('1.0');
    });

    it('should handle corrupted or edge case files gracefully', () => {
      // Test with minimal workbook
      const minimalWb = XLSX.utils.book_new();
      const emptySheet = XLSX.utils.aoa_to_sheet([]);
      XLSX.utils.book_append_sheet(minimalWb, emptySheet, 'Empty');
      
      expect(() => deepCopyWorkbook(minimalWb)).not.toThrow();
      expect(() => detectFileFormatting(minimalWb)).not.toThrow();
      
      // Test with null/undefined
      expect(detectFileFormatting(null as any)).toBe(false);
      expect(detectFileFormatting(undefined as any)).toBe(false);
      
      // Test with corrupted structure
      const corruptedWb = { SheetNames: ['Test'], Sheets: { Test: null } } as any;
      expect(() => deepCopyWorkbook(corruptedWb)).not.toThrow();
    });
  });

  describe('Performance Tests', () => {
    it('should handle large files efficiently', () => {
      // Create a larger test file
      const largeData = [];
      largeData.push(['Name', 'Email', 'Phone', 'Department', 'Position']);
      
      for (let i = 0; i < 1000; i++) {
        largeData.push([
          `User ${i}`,
          `user${i}@example.com`,
          `555-${String(i).padStart(4, '0')}`,
          `Dept ${i % 10}`,
          `Position ${i % 5}`
        ]);
      }
      
      const ws = XLSX.utils.aoa_to_sheet(largeData);
      ws['!cols'] = new Array(5).fill({ width: 15 });
      
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Large');
      
      const startTime = Date.now();
      
      // Process large file
      const copied = deepCopyWorkbook(wb);
      const cloned = deepCloneWorksheet(copied.Sheets['Large']);
      const withColumn = addColumnWithFormatting(cloned, {
        header: 'Status',
        values: new Array(1000).fill('Active')
      });
      
      const endTime = Date.now();
      const processingTime = endTime - startTime;
      
      // Should complete within reasonable time (less than 5 seconds)
      expect(processingTime).toBeLessThan(5000);
      
      // Verify result
      expect(withColumn['F1']?.v).toBe('Status');
      expect(withColumn['F2']?.v).toBe('Active');
      expect(withColumn['!cols']).toHaveLength(6);
    });
  });
});