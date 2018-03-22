package excelProjectW11D2;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReader {
    public static final String SAMPLE_XLSX_FILE_PATH = "170049.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
       	Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

       	//Retrieving the Name of the Workbook
       	System.out.println("Name of the Workbook = 170049.xlsx");
       	
        // Retrieving the number of sheets in the Workbook
       	System.out.println();
        System.out.println();
       	System.out.println("Number of Worksheets that the Workbook have = " + workbook.getNumberOfSheets() + " Worksheets ; ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

     // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Name of each Worksheet:");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("                       -> " + sheet.getSheetName());
        }

        // 2. Or you can use a for-each loop
        System.out.println();
        System.out.println();
        System.out.println("Name of each Worksheet with the number of rows it contains:");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
            
            int rowNum = sheet.getLastRowNum()+1;  
            System.out.println("       ->Number Of Rows '" + sheet.getSheetName() + "' sheet contains = " + rowNum );              
            
        }
            
        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println();
        System.out.println();
        System.out.println("Number of each row with number of columns it contains also with the statement it contains for a particular sheet:");  
        System.out.println();
        System.out.println("For '" + sheet.getSheetName()+"' sheet;");
        System.out.println();
        
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
        
        int rowNum = row.getRowNum();
        
        System.out.println("=> Number of the Row: " + rowNum);
        
        int ColNum = row.getLastCellNum();
        System.out.println("   => Number of the Columns, '"+rowNum +"' Row comtains = " + ColNum);
        System.out.println("      => Columns contain: ");

        // Now let's iterate over the columns of the current row
        Iterator<Cell> cellIterator = row.cellIterator();

        while (cellIterator.hasNext()) {     
            Cell cell = cellIterator.next();
             
            String cellValue = dataFormatter.formatCellValue(cell);
            

            System.out.println ();
            System.out.print("                               "+cellValue + "\t");
        }
        System.out.println();
    }         
               
        // Closing the workbook
        workbook.close();
    }
}