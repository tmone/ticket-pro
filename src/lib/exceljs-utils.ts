/**
 * ExcelJS-based utilities for better formatting preservation
 * To use this, install ExcelJS: npm install exceljs
 */

// Types for ExcelJS (since we might not have @types/exceljs installed)
interface ExcelJSWorkbook {
  xlsx: {
    readFile(filename: string): Promise<void>;
    load(buffer: Buffer): Promise<void>;
    writeFile(filename: string): Promise<void>;
    writeBuffer(): Promise<Buffer>;
  };
  worksheets: ExcelJSWorksheet[];
  getWorksheet(name: string): ExcelJSWorksheet | undefined;
  addWorksheet(name: string): ExcelJSWorksheet;
}

interface ExcelJSWorksheet {
  name: string;
  rowCount: number;
  columnCount: number;
  actualRowCount: number;
  actualColumnCount: number;
  getCell(row: number, col: number): ExcelJSCell;
  getCell(address: string): ExcelJSCell;
  getColumn(col: number): ExcelJSColumn;
  columns: ExcelJSColumn[];
  model: {
    merges?: string[];
  };
}

interface ExcelJSCell {
  value: any;
  type: string;
  font?: any;
  fill?: any;
  border?: any;
  alignment?: any;
  numFmt?: string;
}

interface ExcelJSColumn {
  width?: number;
  hidden?: boolean;
}

/**
 * Enhanced Excel utilities using ExcelJS library
 * Provides much better formatting preservation than XLSX.js
 */
export class ExcelJSHandler {
  private ExcelJS: any;
  private workbook: ExcelJSWorkbook | null = null;
  private originalBuffer: Buffer | null = null;

  constructor() {
    // Try to load ExcelJS
    try {
      this.ExcelJS = require('exceljs');
    } catch (error) {
      throw new Error('ExcelJS not installed. Run: npm install exceljs');
    }
  }

  /**
   * Read Excel file from buffer with full formatting support
   */
  async readFromBuffer(buffer: Buffer): Promise<{
    workbook: ExcelJSWorkbook;
    hasFormatting: boolean;
    sheets: string[];
  }> {
    this.originalBuffer = buffer;
    this.workbook! = new this.ExcelJS.Workbook();
    
    await this.workbook!.xlsx.load(buffer);
    
    return {
      workbook: this.workbook!,
      hasFormatting: this.detectFileFormatting(this.workbook!),
      sheets: this.workbook!.worksheets.map(ws => ws.name)
    };
  }

  /**
   * Detect if file has formatting that should be preserved
   */
  private detectFileFormatting(workbook: ExcelJSWorkbook): boolean {
    try {
      for (const worksheet of workbook.worksheets) {
        // Check for column widths
        if (worksheet.columns.some(col => col.width && col.width !== 10)) {
          return true;
        }

        // Check for merged cells
        if (worksheet.model.merges && worksheet.model.merges.length > 0) {
          return true;
        }

        // Check first few cells for styling
        for (let row = 1; row <= Math.min(5, worksheet.actualRowCount); row++) {
          for (let col = 1; col <= Math.min(5, worksheet.actualColumnCount); col++) {
            const cell = worksheet.getCell(row, col);
            
            if (cell.font || cell.fill || cell.border || cell.alignment) {
              return true;
            }
          }
        }
      }
    } catch (error) {
      console.warn('Error detecting formatting:', error);
      return true; // Assume formatting if we can't detect
    }
    
    return false;
  }

  /**
   * Process sheet data for app consumption (like XLSX.utils.sheet_to_json)
   */
  processSheetData(sheetName: string): {
    headers: string[];
    rows: Record<string, any>[];
  } {
    if (!this.workbook!) {
      throw new Error('No workbook loaded');
    }

    const worksheet = this.workbook!.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    const headers: string[] = [];
    const rows: Record<string, any>[] = [];

    // Get headers from first row
    if (worksheet.actualRowCount > 0) {
      for (let col = 1; col <= worksheet.actualColumnCount; col++) {
        const cell = worksheet.getCell(1, col);
        const cellValue = this.getCellDisplayValue(cell.value);
        headers.push(cellValue || `Column${col}`);
      }
    }

    // Get data rows
    for (let row = 2; row <= worksheet.actualRowCount; row++) {
      const rowData: Record<string, any> = {
        __rowNum__: row,
        checkedInTime: null
      };

      for (let col = 1; col <= worksheet.actualColumnCount; col++) {
        const cell = worksheet.getCell(row, col);
        const header = headers[col - 1];
        if (header) {
          rowData[header] = this.getCellDisplayValue(cell.value);
        }
      }

      rows.push(rowData);
    }

    return { headers, rows };
  }

  /**
   * Add check-in column with full formatting preservation
   */
  async addCheckInColumn(
    sheetName: string, 
    checkInData: { rowNum: number; time: string }[]
  ): Promise<void> {
    if (!this.workbook!) {
      throw new Error('No workbook loaded');
    }

    const worksheet = this.workbook!.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Find the next available column
    const newColNum = worksheet.actualColumnCount + 1;

    // Add header with formatting copied from A1
    const headerCell = worksheet.getCell(1, newColNum);
    headerCell.value = 'Checked-In At';

    // Copy formatting from A1 if it exists
    const a1Cell = worksheet.getCell(1, 1);
    this.copyFormatting(a1Cell, headerCell);

    // Add check-in data
    checkInData.forEach(({ rowNum, time }) => {
      if (rowNum >= 2) { // Skip header row
        const cell = worksheet.getCell(rowNum, newColNum);
        cell.value = time;

        // Copy formatting from corresponding row's first cell
        const refCell = worksheet.getCell(rowNum, 1);
        this.copyFormatting(refCell, cell);
      }
    });

    // Set column width
    const newColumn = worksheet.getColumn(newColNum);
    newColumn.width = 20;
  }

  /**
   * Copy all formatting from source cell to target cell
   */
  private copyFormatting(sourceCell: ExcelJSCell, targetCell: ExcelJSCell): void {
    if (sourceCell.font) {
      targetCell.font = { ...sourceCell.font };
    }
    if (sourceCell.fill) {
      targetCell.fill = { ...sourceCell.fill };
    }
    if (sourceCell.border) {
      targetCell.border = { ...sourceCell.border };
    }
    if (sourceCell.alignment) {
      targetCell.alignment = { ...sourceCell.alignment };
    }
    if (sourceCell.numFmt) {
      targetCell.numFmt = sourceCell.numFmt;
    }
  }

  /**
   * Export modified workbook to buffer
   */
  async exportToBuffer(): Promise<Buffer> {
    if (!this.workbook!) {
      throw new Error('No workbook loaded');
    }

    return await this.workbook!.xlsx.writeBuffer() as Buffer;
  }

  /**
   * Get sheet analysis for debugging
   */
  analyzeSheet(sheetName: string): {
    rowCount: number;
    columnCount: number;
    hasFormatting: boolean;
    formattingDetails: {
      hasColumnWidths: boolean;
      hasMergedCells: boolean;
      hasStyledCells: boolean;
      sampleCellFormatting: any;
    };
  } {
    if (!this.workbook!) {
      throw new Error('No workbook loaded');
    }

    const worksheet = this.workbook!.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Check for column widths
    const hasColumnWidths = worksheet.columns.some(col => col.width);

    // Check for merged cells
    const hasMergedCells = !!(worksheet.model.merges && worksheet.model.merges.length > 0);

    // Check for styled cells
    let hasStyledCells = false;
    let sampleCellFormatting = null;

    for (let row = 1; row <= Math.min(3, worksheet.actualRowCount); row++) {
      for (let col = 1; col <= Math.min(3, worksheet.actualColumnCount); col++) {
        const cell = worksheet.getCell(row, col);
        
        if (cell.font || cell.fill || cell.border || cell.alignment) {
          hasStyledCells = true;
          if (!sampleCellFormatting) {
            sampleCellFormatting = {
              address: `${this.columnToLetter(col)}${row}`,
              value: cell.value,
              font: cell.font,
              fill: cell.fill,
              border: cell.border,
              alignment: cell.alignment
            };
          }
        }
      }
    }

    return {
      rowCount: worksheet.actualRowCount,
      columnCount: worksheet.actualColumnCount,
      hasFormatting: hasColumnWidths || hasMergedCells || hasStyledCells,
      formattingDetails: {
        hasColumnWidths,
        hasMergedCells,
        hasStyledCells,
        sampleCellFormatting
      }
    };
  }

  /**
   * Convert ExcelJS cell value to display string
   */
  private getCellDisplayValue(cellValue: any): string {
    if (cellValue === null || cellValue === undefined) {
      return '';
    }
    
    // Handle different ExcelJS cell value types
    if (typeof cellValue === 'string' || typeof cellValue === 'number') {
      return String(cellValue);
    }
    
    // Handle Date objects
    if (cellValue instanceof Date) {
      return cellValue.toLocaleDateString();
    }
    
    // Handle ExcelJS rich text objects
    if (cellValue && typeof cellValue === 'object') {
      // Check if it's a rich text object
      if (cellValue.richText && Array.isArray(cellValue.richText)) {
        return cellValue.richText.map((rt: any) => rt.text || '').join('');
      }
      
      // Check if it's a formula result
      if (cellValue.result !== undefined) {
        return String(cellValue.result);
      }
      
      // Check if it has a text property
      if (cellValue.text !== undefined) {
        return String(cellValue.text);
      }
      
      // Check if it has a value property
      if (cellValue.value !== undefined) {
        return String(cellValue.value);
      }
      
      // For hyperlinks
      if (cellValue.hyperlink) {
        return cellValue.text || cellValue.hyperlink;
      }
      
      // Fallback to toString if it's a simple object
      try {
        const stringified = JSON.stringify(cellValue);
        if (stringified !== '{}' && stringified !== 'null') {
          return String(cellValue);
        }
      } catch (e) {
        // Ignore stringify errors
      }
    }
    
    return String(cellValue);
  }

  /**
   * Convert column number to letter (1 = A, 2 = B, etc.)
   */
  private columnToLetter(col: number): string {
    let result = '';
    while (col > 0) {
      col--;
      result = String.fromCharCode(65 + (col % 26)) + result;
      col = Math.floor(col / 26);
    }
    return result;
  }
}

/**
 * Factory function to create ExcelJS handler
 * Will fallback to XLSX.js if ExcelJS is not available
 */
export async function createExcelHandler(buffer: Uint8Array): Promise<{
  handler: ExcelJSHandler | null;
  fallbackToXLSX: boolean;
  error?: string;
}> {
  try {
    const handler = new ExcelJSHandler();
    await handler.readFromBuffer(Buffer.from(buffer));
    
    return {
      handler,
      fallbackToXLSX: false
    };
  } catch (error) {
    console.warn('ExcelJS not available, will fallback to XLSX.js:', error);
    
    return {
      handler: null,
      fallbackToXLSX: true,
      error: error instanceof Error ? error.message : 'Unknown error'
    };
  }
}

/**
 * Enhanced version of the main app's handleExport using ExcelJS
 */
export async function handleExportWithExcelJS(
  originalFileData: Uint8Array,
  activeSheetName: string,
  rows: any[]
): Promise<Uint8Array> {
  const handler = new ExcelJSHandler();
  await handler.readFromBuffer(Buffer.from(originalFileData));

  // Prepare check-in data
  const checkInData = rows
    .filter(row => row.checkedInTime)
    .map(row => ({
      rowNum: row.__rowNum__,
      time: new Date(row.checkedInTime).toISOString().replace('T', ' ').substring(0, 19)
    }));

  // Add check-in column
  await handler.addCheckInColumn(activeSheetName, checkInData);

  // Export to buffer
  const buffer = await handler.exportToBuffer();
  return new Uint8Array(buffer);
}