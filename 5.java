import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    public static void main(String[] args) {
        try {
           
            FileInputStream fis = new FileInputStream("sheet1.xlsx"); // Provide the path to your Excel file


            Workbook workbook = new XSSFWorkbook(fis);

          
            Sheet sheet = workbook.getSheetAt(0); 

            
            for (Row row : sheet) {
                for (Cell cell : row) {
                   
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                System.out.print(cell.getDateCellValue() + "\t");
                            } else {
                                System.out.print(cell.getNumericCellValue() + "\t");
                            }
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case BLANK:
                            System.out.print("\t");
                            break;
                        default:
                            System.out.print("[UNKNOWN]\t");
                            break;
                    }
                }
                System.out.println(); // Move to the next row
            }

            // Close the FileInputStream and the workbook
            fis.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
