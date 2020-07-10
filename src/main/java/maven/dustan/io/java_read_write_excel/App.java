package maven.dustan.io.java_read_write_excel;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.LogManager;
import java.util.logging.Logger;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *  A class to read data from an excel file using Apache POI
 */
public class App 
{

	public static final String EXCEL_FILE_PATH = "./readable_writable_excel.xlsx";
	
    public static void main( String[] args ) throws EncryptedDocumentException, InvalidFormatException, IOException {
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(EXCEL_FILE_PATH));


        Logger logger = Logger.getLogger(App.class.getName());

        // Retrieving the number of sheets in the Workbook
        //System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        logger.info("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
        /*
           =============================================================
           Iterating over all the sheets in the workbook
           =============================================================
        */

        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();

        logger.info("Retrieving Sheets using Iterator:");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Creating a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. Obtain a rowIterator and columnIterator and iterate over them
        logger.info("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // iterating over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // Closing the workbook
        workbook.close();
    }
    
    private static void printCellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            case STRING:
                System.out.print(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.print(cell.getDateCellValue());
                } else {
                    System.out.print(cell.getNumericCellValue());
                }
                break;
            case FORMULA:
                System.out.print(cell.getCellFormula());
                break;
            case BLANK:
            default:
                System.out.print("");
                break;
        }

        System.out.print("\t");
    }
}
