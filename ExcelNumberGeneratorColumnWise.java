import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
// Remove conflicting logging dependencies
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class ExcelNumberGeneratorColumnWise {
    private static final Logger logger = LogManager.getLogger(ExcelNumberGeneratorColumnWise.class);

    public static void main(String[] args) {
        long startNumber = 7000000001L;
        long endNumber = 7999999999L;
        int maxRows = 1048576; // Excel has 1,048,576 max rows per sheet
        int maxCols = 26; // A to Z (26 columns)
        String filePath = "GeneratedNumbers.xlsx";

        // Create Workbook
        Workbook workbook = new XSSFWorkbook();
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

            // Move to the next column
            colIndex++;

            // If column limit is reached, move to the next row and reset column index
            if (colIndex >= maxCols) {
                colIndex = 0;
                rowIndex++;
            }

            // If row limit is reached, create a new sheet and reset indexes
            if (rowIndex >= maxRows) {
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
