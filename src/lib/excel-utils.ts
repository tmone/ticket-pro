import * as XLSX from "xlsx";
import type { WorkBook, WorkSheet, CellObject, ExcelDataType } from "xlsx";

/**
 * Creates a deep copy of an Excel workbook while preserving all formatting
 * This is critical for maintaining original file formatting when making modifications
 */
export const deepCopyWorkbook = (workbook: WorkBook): WorkBook => {
  try {
    const newWorkbook = XLSX.utils.book_new();
    
    // Copy all sheets with complete formatting
    Object.keys(workbook.Sheets).forEach(sheetName => {
      try {
        const originalSheet = workbook.Sheets[sheetName];
        const newSheet: WorkSheet = {};
        
        // Copy all cells and sheet properties
        Object.keys(originalSheet).forEach(key => {
          try {
            if (key.startsWith('!')) {
              // Copy sheet properties (merges, cols, rows, etc.)
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
              // Copy cell with all formatting properties
              const cell = originalSheet[key];
              if (cell && typeof cell === 'object') {
                newSheet[key] = JSON.parse(JSON.stringify(cell));
              }
            }
          } catch (error) {
            console.warn(`Error copying cell ${key}:`, error);
            // Fallback: shallow copy
            newSheet[key] = originalSheet[key];
          }
        });
        
        newWorkbook.Sheets[sheetName] = newSheet;
      } catch (error) {
        console.warn(`Error copying sheet ${sheetName}:`, error);
        // Fallback: shallow copy
        newWorkbook.Sheets[sheetName] = workbook.Sheets[sheetName];
      }
    });
    
    // Copy workbook metadata safely
    newWorkbook.SheetNames = [...workbook.SheetNames];
    if (workbook.Props) newWorkbook.Props = JSON.parse(JSON.stringify(workbook.Props));
    if (workbook.Custprops) newWorkbook.Custprops = JSON.parse(JSON.stringify(workbook.Custprops));
    if (workbook.Workbook) newWorkbook.Workbook = JSON.parse(JSON.stringify(workbook.Workbook));
    if (workbook.vbaraw) newWorkbook.vbaraw = workbook.vbaraw;
    
    return newWorkbook;
  } catch (error) {
    console.error('Error in deepCopyWorkbook:', error);
    // Fallback: return original workbook if copy fails
    return workbook;
  }
};

/**
 * Detects if an Excel file has formatting that should be preserved
 * Returns true if any formatting is detected across all sheets
 */
export const detectFileFormatting = (workbook: WorkBook): boolean => {
  try {
    if (!workbook || !workbook.SheetNames) {
      return false;
    }

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      
      if (!sheet || typeof sheet !== 'object') continue;
      
      // Check for sheet-level formatting indicators
      if (sheet['!cols'] || sheet['!rows'] || sheet['!merges'] || 
          sheet['!protect'] || sheet['!autofilter'] || sheet['!margins']) {
        return true;
      }
      
      // Check cells for styles - only check actual cell addresses (like A1, B2, etc.)
      try {
        for (const cellAddress of Object.keys(sheet)) {
          if (!cellAddress.startsWith('!')) {
            try {
              const cell = sheet[cellAddress];
              // Safely check if cell exists and has formatting properties
              if (cell && typeof cell === 'object' && 
                  cell !== null && 
                  !Array.isArray(cell) &&
                  (cell.hasOwnProperty('s') || 
                   cell.hasOwnProperty('z') || 
                   cell.hasOwnProperty('l') || 
                   cell.hasOwnProperty('c'))) {
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
    // If we can't detect formatting, assume it has some
    return true;
  }
  return false;
};

/**
 * Creates a new cell while preserving all formatting from source cell
 * This ensures that when we modify cell values, all styles are maintained
 */
export const preserveCellFormatting = (
  sourceCell: CellObject | undefined, 
  newValue: any, 
  cellType: ExcelDataType = 's'
): CellObject => {
  try {
    const newCell: CellObject = {
      v: newValue,
      t: cellType,
    };
    
    if (sourceCell && typeof sourceCell === 'object') {
      // Preserve all formatting properties with deep copy
      if (sourceCell.s) newCell.s = JSON.parse(JSON.stringify(sourceCell.s)); // Style
      if (sourceCell.z) newCell.z = sourceCell.z; // Number format
      if (sourceCell.l) newCell.l = JSON.parse(JSON.stringify(sourceCell.l)); // Hyperlink
      if (sourceCell.c) newCell.c = JSON.parse(JSON.stringify(sourceCell.c)); // Comments
      if (sourceCell.w) newCell.w = sourceCell.w; // Formatted text
      if (sourceCell.f) newCell.f = sourceCell.f; // Formula
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

/**
 * Enhanced deep clone function for worksheets with better formatting preservation
 */
export const deepCloneWorksheet = (originalWs: WorkSheet): WorkSheet => {
  try {
    const cloned: WorkSheet = {};
    
    // Copy all cell data with complete formatting
    Object.keys(originalWs).forEach(key => {
      if (key.startsWith('!')) {
        // Special worksheet properties
        const value = originalWs[key];
        if (key === '!ref') {
          cloned[key] = value;
        } else if (key === '!cols' && Array.isArray(value)) {
          cloned[key] = value.map(col => col ? JSON.parse(JSON.stringify(col)) : col);
        } else if (key === '!rows' && Array.isArray(value)) {
          cloned[key] = value.map(row => row ? JSON.parse(JSON.stringify(row)) : row);
        } else if (key === '!merges' && Array.isArray(value)) {
          cloned[key] = value.map(merge => JSON.parse(JSON.stringify(merge)));
        } else if (value && typeof value === 'object') {
          // Copy other special properties (autofilter, protect, etc.)
          cloned[key] = JSON.parse(JSON.stringify(value));
        } else {
          cloned[key] = value;
        }
      } else {
        // Regular cell - deep copy with all properties
        const cell = originalWs[key];
        if (cell && typeof cell === 'object') {
          cloned[key] = JSON.parse(JSON.stringify(cell));
        } else {
          cloned[key] = cell;
        }
      }
    });
    
    return cloned;
  } catch (error) {
    console.error('Error in deepCloneWorksheet:', error);
    // Fallback to shallow copy
    return { ...originalWs };
  }
};

/**
 * Advanced function to read Excel file with maximum formatting preservation
 */
export const readExcelWithFormatting = (data: Uint8Array): WorkBook => {
  return XLSX.read(data, {
    type: "array",
    cellStyles: true,
    cellFormula: true,
    cellDates: true,
    cellNF: true,
    sheetStubs: true,
    bookVBA: true,
  });
};

/**
 * Advanced function to write Excel file with maximum formatting preservation
 */
export const writeExcelWithFormatting = (workbook: WorkBook): Uint8Array => {
  return XLSX.write(workbook, { 
    bookType: 'xlsx', 
    type: 'array',
    cellStyles: true,
    cellDates: true,
    bookVBA: true,
    compression: true,
  });
};

/**
 * Safely add a new column to worksheet while preserving all existing formatting
 */
export const addColumnWithFormatting = (
  worksheet: WorkSheet,
  columnData: { header: string; values: string[] },
  preserveFormats = true
): WorkSheet => {
  try {
    // Get current range
    const currentRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
    
    // Calculate new column index
    const newColIndex = currentRange.e.c + 1;
    const newColLetter = XLSX.utils.encode_col(newColIndex);
    
    // Add header for the new column
    const headerCellAddress = `${newColLetter}1`;
    worksheet[headerCellAddress] = {
      v: columnData.header,
      t: 's'
    };
    
    // Add data for each row
    columnData.values.forEach((value, index) => {
      const cellAddress = `${newColLetter}${index + 2}`;
      worksheet[cellAddress] = {
        v: value,
        t: 's'
      };
    });
    
    // Update the worksheet range to include the new column
    const newRange = {
      s: { r: currentRange.s.r, c: currentRange.s.c },
      e: { r: Math.max(currentRange.e.r, columnData.values.length), c: newColIndex }
    };
    worksheet['!ref'] = XLSX.utils.encode_range(newRange);
    
    // Update column widths if they exist
    if (preserveFormats && worksheet['!cols']) {
      const cols = [...worksheet['!cols']];
      // Ensure the array is dense up to the new column
      for (let i = cols.length; i <= newColIndex; i++) {
        if (!cols[i]) {
          cols[i] = { width: 15 }; // Default width
        }
      }
      // Set width for new column
      cols[newColIndex] = { width: 20 };
      worksheet['!cols'] = cols;
    } else if (!worksheet['!cols']) {
      // Create column widths array if it doesn't exist
      const cols = [];
      for (let i = 0; i <= newColIndex; i++) {
        cols[i] = { width: i === newColIndex ? 20 : 15 };
      }
      worksheet['!cols'] = cols;
    }
    
    return worksheet;
  } catch (error) {
    console.error('Error adding column with formatting:', error);
    throw error;
  }
};