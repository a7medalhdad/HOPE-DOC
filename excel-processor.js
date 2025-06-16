const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { PDFDocument: PDFLib } = require('pdf-lib');

// Add XLSX library for better backward compatibility
let XLSX;
try {
  XLSX = require('xlsx');
} catch (error) {
  console.warn('XLSX library not available, falling back to ExcelJS only');
}

// Check if we're running in Electron environment
const isElectron = typeof process !== 'undefined' && process.versions && process.versions.electron;

/**
 * Utility functions for Excel file processing
 */
class ExcelProcessor {
  /**
   * Read an Excel file and return its contents as JSON
   * @param {string} filePath - Path to the Excel file
   * @param {number} sheetIndex - Index of the sheet to read (0-based)
   * @returns {Promise<Object>} - Promise resolving to object with success flag and data
   */
  static async readExcelFile(filePath, sheetIndex = 0) {
    const MAX_RETRIES = 3;
    const RETRY_DELAY_MS = 100;

    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      try {
        console.log(`Reading Excel file: ${filePath} (Attempt ${attempt})`);
        
        // Check if file exists before trying to read
        if (!fs.existsSync(filePath)) {
          throw new Error(`File not found at path: ${filePath}`);
        }
        
        // Check file extension
        const fileExtension = path.extname(filePath).toLowerCase();
        
        // Always use first sheet (index 0) as requested by the user
        const actualSheetIndex = 0;
        
        if (XLSX && (fileExtension === '.xls' || fileExtension === '.xlsx')) {
          try {
            console.log('Attempting to read Excel file using XLSX library');
            const fileData = fs.readFileSync(filePath);
            const workbook = XLSX.read(fileData, { type: 'buffer', cellDates: true, cellStyles: true });
            
            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
              throw new Error('The workbook has no worksheets (XLSX)');
            }
            
            const sheetName = workbook.SheetNames[0];
            console.log(`Using first sheet: '${sheetName}'`);
            const worksheet = workbook.Sheets[sheetName];
            
            if (!worksheet) {
              throw new Error(`First sheet '${sheetName}' is empty or cannot be read`);
            }
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
            console.log(`Successfully read ${jsonData.length} rows from first sheet`);
            
            return {
              success: true,
              data: jsonData,
              filePath: filePath
            };
          } catch (xlsxError) {
            console.warn(`Attempt ${attempt}: XLSX read failed.`, xlsxError.message);
            // Fallback to ExcelJS on failure
          }
        }
        
        // Try with ExcelJS
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        
        let worksheet = workbook.worksheets[0];
        if (!worksheet) {
            throw new Error('The workbook has no worksheets (ExcelJS)');
        }
        console.log(`Using first sheet: '${worksheet.name}'`);
        
        const rows = [];
        worksheet.eachRow((row) => {
          const rowData = [];
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            while (rowData.length < colNumber - 1) {
              rowData.push("");
            }
            let value = cell.value;
            if (value !== null && value !== undefined) {
              if (value.result !== undefined) value = value.result;
              else if (value.text !== undefined) value = value.text;
              else if (value instanceof Date) value = value.toISOString().split('T')[0];
            } else {
              value = "";
            }
            rowData.push(value);
          });
          rows.push(rowData);
        });
        
        return {
          success: true,
          data: rows,
          filePath: filePath
        };
      } catch (error) {
        console.error(`Attempt ${attempt} failed:`, error.message);
        if (attempt === MAX_RETRIES) {
          return { success: false, error: error.message || 'Unknown error occurred while reading Excel file' };
        }
        await new Promise(resolve => setTimeout(resolve, RETRY_DELAY_MS));
      }
    }
  }
  
  /**
   * Write data to an Excel file
   * @param {string} filePath - Path to save the Excel file
   * @param {Array} data - Array of arrays (rows and cells)
   * @returns {Promise<string>} - Promise resolving to the file path
   */
  static async writeExcelFile(filePath, data) {
    try {
      console.log(`Writing Excel file: ${filePath}`);
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      
      worksheet.addRows(data);
      
      await workbook.xlsx.writeFile(filePath);
      return filePath;
    } catch (error) {
      console.error('Error writing Excel file:', error);
      throw error;
    }
  }
  
  /**
   * Process Excel files for ZARA shipment
   * @param {string} referenceFilePath - Path to the reference Excel file
   * @param {Array<string>} dataFilePaths - Paths to the data Excel files
   * @param {string} outputFilePath - Path to save the output Excel file
   * @param {Array<string>} originalFileNames - Original file names for the data files
   * @returns {Promise<Object>} - Processing results
   */
  static async processZaraShipmentFiles(referenceFilePath, dataFilePaths, outputFilePath, originalFileNames = []) {
    try {
      console.log('Starting ZARA shipment file processing');
      console.log(`Reference file: ${referenceFilePath}`);
      console.log(`Data files: ${dataFilePaths.join(', ')}`);
      
      // Read reference file to build lookup table
      let referenceData;
      try {
        const referenceDataResult = await this.readExcelFile(referenceFilePath);
        
        // Handle the data whether it's in the new format (object with data property) or old format (direct array)
        if (referenceDataResult && typeof referenceDataResult === 'object') {
          if (referenceDataResult.data && Array.isArray(referenceDataResult.data)) {
            referenceData = referenceDataResult.data;
          } else if (Array.isArray(referenceDataResult)) {
            referenceData = referenceDataResult;
          }
        } else if (Array.isArray(referenceDataResult)) {
          referenceData = referenceDataResult;
        }
      } catch (error) {
        console.error('Error reading reference file:', error);
        throw new Error(`Error reading reference file: ${error.message}`);
      }
      
      if (!referenceData || !Array.isArray(referenceData)) {
        throw new Error('Reference data is not in the expected format (should be an array)');
      }
      
      console.log(`Read ${referenceData.length} rows from reference file`);
      
      // Build lookup table
      const lookupTable = new Map();
      
      for (let i = 0; i < referenceData.length; i++) {
        const row = referenceData[i];
        if (row && row[0] != null && row[1] != null) {
          const productName = String(row[0]).trim().toUpperCase();
          const hsCode = String(row[1]);
          if (productName) lookupTable.set(productName, hsCode);
        }
      }
      
      console.log(`Lookup table built with ${lookupTable.size} entries`);
      
      // Process each data file
      const allProcessedData = [];
      const headerRow = ['OLD HSCODE', 'NEW HSCODE', 'COO', 'DES', 'QTY', 'AMOUNT', 'shipmentnumber'];
      allProcessedData.push(headerRow);
      
      const summary = {
        processedFiles: 0,
        totalItemsProcessed: 0,
        itemsNotFound: []
      };
      
      console.log(`Total data files to process: ${dataFilePaths.length}`);
      console.log(`Original file names: ${JSON.stringify(originalFileNames)}`);
      
      for (let i = 0; i < dataFilePaths.length; i++) {
        const dataFilePath = dataFilePaths[i];
        // Use original file name if available, otherwise use the temp file name
        const tempFileName = path.basename(dataFilePath);
        const originalFileName = (originalFileNames && i < originalFileNames.length) ? originalFileNames[i] : tempFileName;
        console.log(`Processing file ${i+1}/${dataFilePaths.length}: ${originalFileName} (temp: ${tempFileName})`);
        
        // Read data file
        let dataFileContent;
        try {
          const dataFileResult = await this.readExcelFile(dataFilePath);
          
          // Handle the data whether it's in the new format or old format
          if (dataFileResult && typeof dataFileResult === 'object') {
            if (dataFileResult.data && Array.isArray(dataFileResult.data)) {
              dataFileContent = dataFileResult.data;
            } else if (Array.isArray(dataFileResult)) {
              dataFileContent = dataFileResult;
            }
          } else if (Array.isArray(dataFileResult)) {
            dataFileContent = dataFileResult;
          }
        } catch (error) {
          console.error(`Error reading data file ${tempFileName}:`, error);
          throw new Error(`Error reading data file ${tempFileName}: ${error.message}`);
        }
        
        if (!dataFileContent || !Array.isArray(dataFileContent)) {
          throw new Error(`Data file ${tempFileName} is not in the expected format (should be an array)`);
        }
        
        console.log(`Read ${dataFileContent.length} rows from data file ${tempFileName} (original: ${originalFileName})`);
        
        let workingData = [...dataFileContent];
        
        // Skip header rows if necessary (first 7 rows for ZARA data)
        if (workingData.length >= 7) {
          workingData = workingData.slice(7);
        }
        
        // Delete columns that aren't needed (original VLOOKUP behavior)
        const deleteColumn = (data, columnIndex) => {
          for (const row of data) {
            if (row && Array.isArray(row) && row.length > columnIndex) {
              row.splice(columnIndex, 1);
            }
          }
        };
        
        // Apply the same column transformations as before
        deleteColumn(workingData, 1); // Delete column B (index 1)
        deleteColumn(workingData, 3); // Delete what was column E, now is D after first deletion (index 3)
        if (workingData.length > 0 && workingData[0] && workingData[0].length > 5) {
          deleteColumn(workingData, 5); // Delete column G if it exists
        }
        
        // Process data rows - VLOOKUP style
        for (const row of workingData) {
          if (!row || row.length < 2) continue;
          
          // Get the lookup value (product name) from column B (index 1) after deletions
          const productNameForLookup = row[1] ? String(row[1]).trim().toUpperCase() : null;
          
          if (!productNameForLookup) continue;
          
          // Get other values
          const oldHscode = row[0] || '';
          const coo = row[2] || '';
          const qty = row[3] || '';
          const amount = row[4] || '';
          
          // Perform the VLOOKUP
          const newHsCode = lookupTable.get(productNameForLookup) || '';
          
          // Use original file name without extension for shipmentnumber
          // First clean up the name to make sure it's valid
          let shipmentNumber = originalFileName;
          
          // Remove file extension if present
          shipmentNumber = shipmentNumber.replace(/\.[^\.]+$/, '');
          
          // Make sure the name is clean and usable
          shipmentNumber = shipmentNumber.trim();
          
          allProcessedData.push([
            oldHscode,          // OLD HSCODE (first column)
            newHsCode,          // NEW HSCODE (VLOOKUP result)
            coo,                // COO 
            productNameForLookup, // DES (the lookup key)
            qty,                // QTY
            amount,             // AMOUNT
            shipmentNumber      // shipmentnumber from original filename
          ]);
          
          if (newHsCode) {
            summary.totalItemsProcessed++;
          } else {
            summary.itemsNotFound.push(productNameForLookup);
          }
        }
        
        summary.processedFiles++;
      }
      
      // Write results to output file
      await this.writeExcelFile(outputFilePath, allProcessedData);
      
      console.log('Processing completed successfully');
      console.log(`Output saved to: ${outputFilePath}`);
      
      return {
        success: true,
        summary,
        outputFilePath
      };
    } catch (error) {
      console.error('Error processing ZARA shipment files:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Convert Excel files to PDF
   * @param {Array<string>} excelFilePaths - Paths to the Excel files
   * @param {string} outputPdfPath - Path to save the output PDF file
   * @returns {Promise<Object>} - Processing results
   */
  static async convertExcelToPdf(excelFilePaths, outputPdfPath) {
    try {
      console.log(`Converting ${excelFilePaths.length} Excel files to PDF`);
      console.log(`Output PDF: ${outputPdfPath}`);
      
      // Create a PDF document
      const doc = new PDFDocument({ autoFirstPage: false, layout: 'landscape' });
      const stream = fs.createWriteStream(outputPdfPath);
      doc.pipe(stream);

      let pageCount = 0;
      
      // Process each Excel file
      for (let i = 0; i < excelFilePaths.length; i++) {
        const excelFilePath = excelFilePaths[i];
        const fileName = path.basename(excelFilePath);
        console.log(`Processing Excel file ${i+1}/${excelFilePaths.length}: ${fileName}`);
        
        // Read the Excel file
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        
        // Process each worksheet
        for (const worksheet of workbook.worksheets) {
          // Add a new page for each worksheet
          doc.addPage();
          pageCount++;
          
          // Add file and sheet name at the top
          doc.fontSize(14).text(fileName, 40, 40);
          doc.fontSize(12).text(`Sheet: ${worksheet.name}`, 40, 60);
          
          // Set starting position for the table
          const startX = 40;
          const startY = 90;
          const cellPadding = 5;
          let currentY = startY;
          
          // Calculate column widths
          const colWidths = [];
          const maxCols = Math.min(worksheet.columnCount || 10, 15); // Limit to 15 columns max for readability
          
          for (let col = 0; col < maxCols; col++) {
            colWidths.push(80); // Default width
          }
          
          // Draw headers
          let currentX = startX;
          doc.fontSize(10).fillColor('#000000');
          for (let col = 0; col < maxCols; col++) {
            const cell = worksheet.getRow(1).getCell(col + 1);
            const value = cell.value ? cell.value.toString() : '';
            
            // Draw cell background and border
            doc.rect(currentX, currentY, colWidths[col], 20)
               .fillAndStroke('#e0e0e0', '#000000');
            
            // Draw text
            doc.fillColor('#000000')
               .text(value, currentX + cellPadding, currentY + cellPadding, {
                 width: colWidths[col] - (2 * cellPadding),
                 height: 20 - (2 * cellPadding),
                 align: 'left'
               });
            
            currentX += colWidths[col];
          }
          currentY += 20;
          
          // Draw data rows
          const maxRows = Math.min(worksheet.rowCount, 50); // Limit to 50 rows max per sheet
          for (let row = 2; row <= maxRows; row++) {
            currentX = startX;
            const rowHeight = 18;
            
            for (let col = 0; col < maxCols; col++) {
              const cell = worksheet.getRow(row).getCell(col + 1);
              let value = '';
              
              if (cell.value !== null && cell.value !== undefined) {
                if (cell.value instanceof Date) {
                  value = cell.value.toLocaleDateString();
                } else if (typeof cell.value === 'object' && cell.value.formula) {
                  value = cell.value.result || '';
                } else {
                  value = cell.value.toString();
                }
              }
              
              // Draw cell border
              doc.rect(currentX, currentY, colWidths[col], rowHeight)
                 .stroke('#cccccc');
              
              // Draw text
              doc.fillColor('#000000')
                 .text(value, currentX + cellPadding, currentY + cellPadding, {
                   width: colWidths[col] - (2 * cellPadding),
                   height: rowHeight - (2 * cellPadding),
                   align: 'left'
                 });
              
              currentX += colWidths[col];
            }
            
            currentY += rowHeight;
            
            // Check if we need a new page
            if (currentY > doc.page.height - 40) {
              doc.addPage();
              pageCount++;
              currentY = startY;
              
              // Repeat headers on new page
              currentX = startX;
              for (let col = 0; col < maxCols; col++) {
                const cell = worksheet.getRow(1).getCell(col + 1);
                const value = cell.value ? cell.value.toString() : '';
                
                doc.rect(currentX, currentY, colWidths[col], 20)
                   .fillAndStroke('#e0e0e0', '#000000');
                
                doc.fillColor('#000000')
                   .text(value, currentX + cellPadding, currentY + cellPadding, {
                     width: colWidths[col] - (2 * cellPadding),
                     height: 20 - (2 * cellPadding),
                     align: 'left'
                   });
                
                currentX += colWidths[col];
              }
              currentY += 20;
            }
          }
        }
      }
      
      // Finalize the PDF
      doc.end();
      
      // Wait for the stream to finish
      await new Promise((resolve, reject) => {
        stream.on('finish', resolve);
        stream.on('error', reject);
      });
      
      console.log(`PDF conversion completed successfully with ${pageCount} pages`);
      
      return {
        success: true,
        outputPath: outputPdfPath,
        pageCount,
        message: `Successfully converted ${excelFilePaths.length} Excel files to PDF with ${pageCount} pages`
      };
    } catch (error) {
      console.error('Error converting Excel to PDF:', error);
      return {
        success: false,
        error: error.message || 'Unknown error occurred during Excel to PDF conversion'
      };
    }
  }
  
  /**
   * Write data tables to an Excel file
   * @param {Array<Array<Array<string>>>} tables - Array of tables, where each table is an array of rows, and each row is an array of cells
   * @param {string} outputFilePath - Path to save the output Excel file
   * @returns {Promise<Object>} - Processing results
   */
  static async writeTablesFile(tables, outputFilePath) {
    try {
      console.log(`Writing ${tables.length} tables to Excel file: ${outputFilePath}`);
      const workbook = new ExcelJS.Workbook();
      
      tables.forEach((tableData, index) => {
        if (tableData && tableData.length > 0) {
          const worksheet = workbook.addWorksheet(`Table ${index + 1}`);
          
          // Add data to worksheet
          tableData.forEach(row => {
            worksheet.addRow(row);
          });
          
          /*
          // Auto-fit columns width based on content - THIS IS VERY SLOW
          worksheet.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
              const columnLength = cell.value ? String(cell.value).length : 10;
              if (columnLength > maxLength) {
                maxLength = columnLength;
              }
            });
            column.width = maxLength < 10 ? 10 : maxLength + 2;
          });
          */
        }
      });
      
      // Save the workbook
      await workbook.xlsx.writeFile(outputFilePath);
      
      return {
        success: true,
        outputFilePath: outputFilePath
      };
    } catch (error) {
      console.error('Error writing tables to Excel file:', error);
      return {
        success: false,
        error: error.message || 'Unknown error occurred while writing Excel file'
      };
    }
  }
  
  /**
   * Merge multiple Excel files into a single Excel file
   * @param {Array<string>} excelFilePaths - Paths to the Excel files
   * @param {string} outputExcelPath - Path to save the output Excel file
   * @param {Object} options - Merge options
   * @returns {Promise<Object>} - Processing results
   */
  static async mergeExcelFiles(excelFilePaths, outputExcelPath, options = {}) {
    try {
      console.log(`Merging ${excelFilePaths.length} Excel files`);
      console.log(`Output Excel: ${outputExcelPath}`);
      
      const { includeFileName = true, mergeMode = 'sheets' } = options;
      
      // Create a new workbook for the output
      const outputWorkbook = new ExcelJS.Workbook();
      
      if (mergeMode === 'sheets') {
        // Each input file becomes a separate worksheet in the output file
        for (let i = 0; i < excelFilePaths.length; i++) {
          const excelFilePath = excelFilePaths[i];
          const fileName = path.basename(excelFilePath, path.extname(excelFilePath));
          console.log(`Processing file ${i+1}/${excelFilePaths.length}: ${fileName}`);
          
          // Read the input workbook
          const inputWorkbook = new ExcelJS.Workbook();
          await inputWorkbook.xlsx.readFile(excelFilePath);
          
          // Process each worksheet
          for (const worksheet of inputWorkbook.worksheets) {
            const sheetName = includeFileName ? 
              `${fileName}_${worksheet.name}`.substring(0, 31) : // Excel has a 31 char limit for sheet names
              worksheet.name;
            
            // Create a new worksheet in the output workbook
            const outputWorksheet = outputWorkbook.addWorksheet(sheetName);
            
            // Copy all rows from the input worksheet to the output worksheet
            worksheet.eachRow((row, rowNumber) => {
              const outputRow = outputWorksheet.getRow(rowNumber);
              
              row.eachCell((cell, colNumber) => {
                const outputCell = outputRow.getCell(colNumber);
                outputCell.value = cell.value;
                outputCell.style = cell.style;
              });
              
              outputRow.commit();
            });
            
            // Copy column properties
            worksheet.columns.forEach((col, index) => {
              if (col.width) {
                outputWorksheet.getColumn(index + 1).width = col.width;
              }
            });
          }
        }
      } else if (mergeMode === 'rows') {
        // All input data is merged into a single worksheet
        const outputWorksheet = outputWorkbook.addWorksheet('Merged Data');
        let currentRow = 1;
        
        for (let i = 0; i < excelFilePaths.length; i++) {
          const excelFilePath = excelFilePaths[i];
          const fileName = path.basename(excelFilePath);
          console.log(`Processing file ${i+1}/${excelFilePaths.length}: ${fileName}`);
          
          // Read the input workbook
          const inputWorkbook = new ExcelJS.Workbook();
          await inputWorkbook.xlsx.readFile(excelFilePath);
          
          // We'll use the first worksheet from each file
          const worksheet = inputWorkbook.worksheets[0];
          
          // Add a file name row if requested
          if (includeFileName) {
            const fileNameRow = outputWorksheet.getRow(currentRow++);
            fileNameRow.getCell(1).value = `File: ${fileName}`;
            fileNameRow.font = { bold: true };
            fileNameRow.commit();
          }
          
          // Copy all rows from the input worksheet to the output worksheet
          worksheet.eachRow((row, rowNumber) => {
            const outputRow = outputWorksheet.getRow(currentRow++);
            
            row.eachCell((cell, colNumber) => {
              const outputCell = outputRow.getCell(colNumber);
              outputCell.value = cell.value;
              outputCell.style = cell.style;
            });
            
            outputRow.commit();
          });
          
          // Add a blank row between files
          currentRow++;
        }
      }
      
      // Write the output workbook to disk
      await outputWorkbook.xlsx.writeFile(outputExcelPath);
      
      console.log('Excel merge completed successfully');
      
      return {
        success: true,
        outputPath: outputExcelPath,
        message: `Successfully merged ${excelFilePaths.length} Excel files`
      };
    } catch (error) {
      console.error('Error merging Excel files:', error);
      return {
        success: false,
        error: error.message || 'Unknown error occurred during Excel merge'
      };
    }
  }

  /**
   * Write JSON data to an Excel file
   * @param {Array} jsonData - Array of objects to write to Excel
   * @param {string} outputFilePath - Path to save the Excel file
   * @returns {Promise<Object>} - Promise resolving to object with success flag
   */
  static async writeJsonToExcel(jsonData, outputFilePath) {
    try {
      console.log(`Writing JSON data to Excel file: ${outputFilePath}`);
      
      if (!Array.isArray(jsonData)) {
        throw new Error('JSON data must be an array');
      }
      
      // Create a new workbook
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('بيانات_مستخرجة');
      
      // Add headers if data exists
      if (jsonData.length > 0) {
        // Get headers from the first object's keys
        const headers = Object.keys(jsonData[0]);
        
        // Add header row
        worksheet.addRow(headers);
        
        // Add data rows
        jsonData.forEach(item => {
          const rowValues = headers.map(header => item[header] || '');
          worksheet.addRow(rowValues);
        });
        
        // Format headers
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        
        // Auto-fit columns
        headers.forEach((header, i) => {
          const column = worksheet.getColumn(i + 1);
          let maxLength = header.length;
          
          // Find the maximum content length in the column
          jsonData.forEach(item => {
            const value = String(item[header] || '');
            if (value.length > maxLength) {
              maxLength = value.length;
            }
          });
          
          // Set column width (with some padding)
          column.width = maxLength + 3;
        });
      }
      
      // Write the file
      await workbook.xlsx.writeFile(outputFilePath);
      
      return {
        success: true,
        outputFilePath
      };
    } catch (error) {
      console.error(`Error writing JSON to Excel: ${error.message}`);
      return {
        success: false,
        error: error.message
      };
    }
  }
}

// Make the processor available for command-line usage
if (require.main === module) {
  const args = process.argv.slice(2);
  const command = args[0];
  
  if (command === 'process-zara') {
    const referenceFilePath = args[1];
    const dataFilePaths = args[2].split(',');
    const outputFilePath = args[3];
    
    ExcelProcessor.processZaraShipmentFiles(referenceFilePath, dataFilePaths, outputFilePath)
      .then(result => {
        console.log(JSON.stringify(result));
        process.exit(0);
      })
      .catch(error => {
        console.error(error);
        process.exit(1);
      });
  } else if (command === 'convert-to-pdf') {
    const excelFilePaths = args[1].split(',');
    const outputPdfPath = args[2];
    
    ExcelProcessor.convertExcelToPdf(excelFilePaths, outputPdfPath)
      .then(result => {
        console.log(JSON.stringify(result));
        process.exit(0);
      })
      .catch(error => {
        console.error(error);
        process.exit(1);
      });
  } else if (command === 'merge-excel') {
    const excelFilePaths = args[1].split(',');
    const outputExcelPath = args[2];
    const options = args[3] ? JSON.parse(args[3]) : {};
    
    ExcelProcessor.mergeExcelFiles(excelFilePaths, outputExcelPath, options)
      .then(result => {
        console.log(JSON.stringify(result));
        process.exit(0);
      })
      .catch(error => {
        console.error(error);
        process.exit(1);
      });
  } else if (command === 'json-to-excel') {
    const jsonDataPath = args[1];
    const outputExcelPath = args[2];
    
    // Read the JSON data from file
    const jsonData = JSON.parse(fs.readFileSync(jsonDataPath, 'utf8'));
    
    ExcelProcessor.writeJsonToExcel(jsonData, outputExcelPath)
      .then(result => {
        console.log(JSON.stringify(result));
        process.exit(0);
      })
      .catch(error => {
        console.error(error);
        process.exit(1);
      });
  } else {
    console.error('Unknown command:', command);
    console.error('Available commands: process-zara, convert-to-pdf, merge-excel, json-to-excel');
    process.exit(1);
  }
}

module.exports = ExcelProcessor;
