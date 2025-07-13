/**
 * @jest-environment jsdom
 */
import * as XLSX from 'xlsx';

// Mock data for testing
const createMockWorkbook = (withFormatting = true): XLSX.WorkBook => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Name', 'Email', 'Phone'],
    ['John Doe', 'john@example.com', '123-456-7890'],
    ['Jane Smith', 'jane@example.com', '098-765-4321']
  ]);

  if (withFormatting) {
    // Add some mock formatting
    ws['!cols'] = [
      { width: 15 },
      { width: 25 },
      { width: 20 }
    ];
    
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }
    ];

    // Add cell formatting
    if (ws['A1']) {
      ws['A1'].s = {
        font: { bold: true, color: { rgb: 'FF0000' } },
        fill: { fgColor: { rgb: 'FFFF00' } }
      };
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return wb;
};

// Helper functions to test
const deepCopyWorkbook = (workbook: XLSX.WorkBook): XLSX.WorkBook => {
  try {
    const newWorkbook = XLSX.utils.book_new();
    
    Object.keys(workbook.Sheets).forEach(sheetName => {
      try {
        const originalSheet = workbook.Sheets[sheetName];
        const newSheet: XLSX.WorkSheet = {};
        
        Object.keys(originalSheet).forEach(key => {
          try {
            if (key.startsWith('!')) {
              newSheet[key] = Array.isArray(originalSheet[key]) 
                ? [...originalSheet[key]] 
                : { ...originalSheet[key] };
            } else {
              const cell = originalSheet[key];
              if (cell && typeof cell === 'object') {
                newSheet[key] = { ...cell };
              }
            }
          } catch (error) {
            console.warn(`Error copying cell ${key}:`, error);
          }
        });
        
        newWorkbook.Sheets[sheetName] = newSheet;
      } catch (error) {
        console.warn(`Error copying sheet ${sheetName}:`, error);
      }
    });
    
    newWorkbook.SheetNames = [...workbook.SheetNames];
    if (workbook.Props) newWorkbook.Props = { ...workbook.Props };
    if (workbook.Custprops) newWorkbook.Custprops = { ...workbook.Custprops };
    if (workbook.Workbook) newWorkbook.Workbook = { ...workbook.Workbook };
    if (workbook.vbaraw) newWorkbook.vbaraw = workbook.vbaraw;
    
    return newWorkbook;
  } catch (error) {
    console.error('Error in deepCopyWorkbook:', error);
    return workbook;
  }
};

const detectFileFormatting = (workbook: XLSX.WorkBook): boolean => {
  try {
    if (!workbook || !workbook.SheetNames) {
      return false;
    }

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      
      if (!sheet || typeof sheet !== 'object') continue;
      
      if (sheet['!cols'] || sheet['!rows'] || sheet['!merges']) {
        return true;
      }
      
      try {
        for (const cellAddress of Object.keys(sheet)) {
          if (!cellAddress.startsWith('!')) {
            try {
              const cell = sheet[cellAddress];
              if (cell && typeof cell === 'object' && 
                  cell !== null && 
                  !Array.isArray(cell) &&
                  (cell.hasOwnProperty('s') || cell.hasOwnProperty('z'))) {
                return true;
              }
            } catch (cellError) {
              console.warn(`Error checking cell ${cellAddress}:`, cellError);
              continue;
            }
          }
        }
      } catch (sheetError) {
        console.warn(`Error checking sheet ${sheetName}:`, sheetError);
        continue;
      }
    }
  } catch (error) {
    console.warn('Error detecting file formatting:', error);
    return true;
  }
  return false;
};

const preserveCellFormatting = (sourceCell: XLSX.CellObject | undefined, newValue: any, cellType: XLSX.ExcelDataType = 's') => {
  try {
    const newCell: XLSX.CellObject = {
      v: newValue,
      t: cellType,
    };
    
    if (sourceCell && typeof sourceCell === 'object') {
      if (sourceCell.s) newCell.s = sourceCell.s;
      if (sourceCell.z) newCell.z = sourceCell.z;
      if (sourceCell.l) newCell.l = sourceCell.l;
      if (sourceCell.c) newCell.c = sourceCell.c;
      if (sourceCell.w) newCell.w = sourceCell.w;
    }
    
    return newCell;
  } catch (error) {
    console.warn('Error preserving cell formatting:', error);
    return {
      v: newValue,
      t: cellType,
    };
  }
};

describe('Excel Utilities', () => {
  describe('deepCopyWorkbook', () => {
    it('should create a deep copy of workbook while preserving structure', () => {
      const originalWb = createMockWorkbook(true);
      const copiedWb = deepCopyWorkbook(originalWb);

      expect(copiedWb).not.toBe(originalWb);
      expect(copiedWb.SheetNames).toEqual(originalWb.SheetNames);
      expect(copiedWb.SheetNames).not.toBe(originalWb.SheetNames);
    });

    it('should preserve sheet formatting in copied workbook', () => {
      const originalWb = createMockWorkbook(true);
      const copiedWb = deepCopyWorkbook(originalWb);

      const originalSheet = originalWb.Sheets['Sheet1'];
      const copiedSheet = copiedWb.Sheets['Sheet1'];

      expect(copiedSheet['!cols']).toBeDefined();
      expect(copiedSheet['!cols']).toEqual(originalSheet['!cols']);
      expect(copiedSheet['!cols']).not.toBe(originalSheet['!cols']);
    });

    it('should preserve cell formatting in copied workbook', () => {
      const originalWb = createMockWorkbook(true);
      const copiedWb = deepCopyWorkbook(originalWb);

      const originalCell = originalWb.Sheets['Sheet1']['A1'];
      const copiedCell = copiedWb.Sheets['Sheet1']['A1'];

      if (originalCell?.s) {
        expect(copiedCell?.s).toBeDefined();
        expect(copiedCell?.s).toEqual(originalCell.s);
        expect(copiedCell?.s).not.toBe(originalCell.s);
      }
    });

    it('should handle workbook without formatting gracefully', () => {
      const originalWb = createMockWorkbook(false);
      const copiedWb = deepCopyWorkbook(originalWb);

      expect(copiedWb).not.toBe(originalWb);
      expect(copiedWb.SheetNames).toEqual(originalWb.SheetNames);
    });

    it('should handle corrupted workbook gracefully', () => {
      const corruptedWb = { SheetNames: ['Sheet1'], Sheets: { Sheet1: null } } as any;
      const result = deepCopyWorkbook(corruptedWb);

      expect(result).toBe(corruptedWb); // Should return original if copy fails
    });
  });

  describe('detectFileFormatting', () => {
    it('should detect formatting when columns are defined', () => {
      const wb = createMockWorkbook(true);
      const hasFormatting = detectFileFormatting(wb);
      expect(hasFormatting).toBe(true);
    });

    it('should detect formatting when merges are defined', () => {
      const wb = createMockWorkbook(false);
      wb.Sheets['Sheet1']['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
      
      const hasFormatting = detectFileFormatting(wb);
      expect(hasFormatting).toBe(true);
    });

    it('should detect formatting when cells have style properties', () => {
      const wb = createMockWorkbook(false);
      wb.Sheets['Sheet1']['A1'] = {
        v: 'Test',
        t: 's',
        s: { font: { bold: true } }
      };
      
      const hasFormatting = detectFileFormatting(wb);
      expect(hasFormatting).toBe(true);
    });

    it('should return false for workbook without formatting', () => {
      const wb = createMockWorkbook(false);
      delete wb.Sheets['Sheet1']['!cols'];
      
      const hasFormatting = detectFileFormatting(wb);
      expect(hasFormatting).toBe(false);
    });

    it('should handle null/undefined workbook gracefully', () => {
      expect(detectFileFormatting(null as any)).toBe(false);
      expect(detectFileFormatting(undefined as any)).toBe(false);
      expect(detectFileFormatting({} as any)).toBe(false);
    });

    it('should handle corrupted sheet gracefully', () => {
      const wb = {
        SheetNames: ['Sheet1'],
        Sheets: { Sheet1: 'invalid' }
      } as any;
      
      const hasFormatting = detectFileFormatting(wb);
      expect(hasFormatting).toBe(true); // Should assume formatting if can't detect
    });
  });

  describe('preserveCellFormatting', () => {
    it('should preserve style properties from source cell', () => {
      const sourceCell: XLSX.CellObject = {
        v: 'Original',
        t: 's',
        s: { font: { bold: true, color: { rgb: 'FF0000' } } },
        z: '@'
      };

      const newCell = preserveCellFormatting(sourceCell, 'New Value', 's');

      expect(newCell.v).toBe('New Value');
      expect(newCell.t).toBe('s');
      expect(newCell.s).toEqual(sourceCell.s);
      expect(newCell.z).toBe(sourceCell.z);
    });

    it('should preserve all formatting properties', () => {
      const sourceCell: XLSX.CellObject = {
        v: 'Original',
        t: 's',
        s: { font: { bold: true } },
        z: '@',
        l: { Target: 'http://example.com' },
        c: [{ a: 'Author', t: 'Comment' }],
        w: 'Formatted Text'
      };

      const newCell = preserveCellFormatting(sourceCell, 'New Value', 's');

      expect(newCell.s).toEqual(sourceCell.s);
      expect(newCell.z).toBe(sourceCell.z);
      expect(newCell.l).toEqual(sourceCell.l);
      expect(newCell.c).toEqual(sourceCell.c);
      expect(newCell.w).toBe(sourceCell.w);
    });

    it('should create basic cell when no source formatting exists', () => {
      const newCell = preserveCellFormatting(undefined, 'New Value', 'n');

      expect(newCell.v).toBe('New Value');
      expect(newCell.t).toBe('n');
      expect(newCell.s).toBeUndefined();
      expect(newCell.z).toBeUndefined();
    });

    it('should handle null source cell gracefully', () => {
      const newCell = preserveCellFormatting(null as any, 'New Value', 's');

      expect(newCell.v).toBe('New Value');
      expect(newCell.t).toBe('s');
    });

    it('should handle corrupted source cell gracefully', () => {
      const corruptedCell = 'not an object' as any;
      const newCell = preserveCellFormatting(corruptedCell, 'New Value', 's');

      expect(newCell.v).toBe('New Value');
      expect(newCell.t).toBe('s');
    });
  });

  describe('Excel Integration Tests', () => {
    it('should maintain formatting through full read-modify-write cycle', () => {
      // Create a formatted workbook
      const originalWb = createMockWorkbook(true);
      
      // Convert to buffer (simulate file upload)
      const buffer = XLSX.write(originalWb, { 
        type: 'array', 
        bookType: 'xlsx',
        cellStyles: true 
      });
      
      // Read back with formatting
      const readWb = XLSX.read(buffer, {
        type: 'array',
        cellStyles: true,
        cellFormula: true,
        cellDates: true,
        cellNF: true
      });

      // Verify formatting is preserved
      const hasFormatting = detectFileFormatting(readWb);
      expect(hasFormatting).toBe(true);

      // Make a copy and modify
      const modifiedWb = deepCopyWorkbook(readWb);
      const sheet = modifiedWb.Sheets['Sheet1'];
      
      // Add a new column with data
      sheet['D1'] = { v: 'Check-in Time', t: 's' };
      sheet['D2'] = { v: '2024-01-01 10:00', t: 's' };
      
      // Update range
      const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:C3');
      range.e.c = 3; // Extend to column D
      sheet['!ref'] = XLSX.utils.encode_range(range);

      // Write back to buffer
      const finalBuffer = XLSX.write(modifiedWb, {
        type: 'array',
        bookType: 'xlsx',
        cellStyles: true
      });

      // Verify the final result
      const finalWb = XLSX.read(finalBuffer, {
        type: 'array',
        cellStyles: true
      });

      expect(detectFileFormatting(finalWb)).toBe(true);
      expect(finalWb.Sheets['Sheet1']['D1']?.v).toBe('Check-in Time');
      expect(finalWb.Sheets['Sheet1']['D2']?.v).toBe('2024-01-01 10:00');
    });

    it('should handle multiple sheets with different formatting', () => {
      const wb = XLSX.utils.book_new();
      
      // Create two sheets with different formatting
      const sheet1 = XLSX.utils.aoa_to_sheet([['Data1'], ['Value1']]);
      sheet1['!cols'] = [{ width: 15 }];
      
      const sheet2 = XLSX.utils.aoa_to_sheet([['Data2'], ['Value2']]);
      sheet2['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
      
      XLSX.utils.book_append_sheet(wb, sheet1, 'FormattedSheet1');
      XLSX.utils.book_append_sheet(wb, sheet2, 'FormattedSheet2');

      const copiedWb = deepCopyWorkbook(wb);

      expect(copiedWb.Sheets['FormattedSheet1']['!cols']).toBeDefined();
      expect(copiedWb.Sheets['FormattedSheet2']['!merges']).toBeDefined();
      expect(detectFileFormatting(copiedWb)).toBe(true);
    });
  });
});