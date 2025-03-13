import org.apache.poi.ss.usermodel.*;
// import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class ExcelNumberGeneratorColumnWise {
    private static final Logger logger = LogManager.getLogger(ExcelNumberGeneratorColumnWise.class);

    public static void main(String[] args) {
        long startNumber = 7000000001L;
        long endNumber = 7000000100L;
        int maxRows = 1048576; // Excel has 1,048,576 max rows per sheet
        int maxCols = 16384; // A to Z (26 columns)
        String filePath = "GeneratedNumbers.xlsx";

        // Create Workbook
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        // Workbook workbook = new XSSFWorkbook();
        int sheetNumber = 1;
        Sheet sheet = workbook.createSheet("Sheet_" + sheetNumber);

        long currentNumber = startNumber;
        int rowIndex = 0;
        int colIndex = 0;

        while (currentNumber <= endNumber) {
            // Create new row if first column
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            // Create cell and add value
            Cell cell = row.createCell(colIndex);
            cell.setCellValue(currentNumber);
            currentNumber++;

            // Move to the next row
            rowIndex++;

            // If row limit is reached, move to the next column and reset row index
            if (rowIndex >= maxRows) {
                rowIndex = 0;
                colIndex++;
            }

            // If column limit is reached, create a new sheet and reset indexes
            if (colIndex >= maxCols) {
                sheetNumber++;
                sheet = workbook.createSheet("Sheet_" + sheetNumber);
                rowIndex = 0;
                colIndex = 0;
            }

            // Print progress every 10 million numbers
            if (currentNumber % 10000000 == 0) {
                logger.info("Generated up to: " + currentNumber);
            }
        }

        // Write to file
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
            logger.info("Excel file created successfully: " + filePath);
        } catch (IOException e) {
            logger.error("Error writing Excel file", e);
        }

        // Close workbook
        try {
            workbook.close();
        } catch (IOException e) {
            logger.error("Error closing workbook", e);
        }
    }
}
