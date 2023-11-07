import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {
    public static void main(String[] args) {
        
        Workbook workbook = new XSSFWorkbook(); 
       
        Sheet sheet = workbook.createSheet("Sheet1");


        String[] headers = {"Name", "Age", "Email"};
        Object[][] data = {
            {"John Doe", 30, "john@test.com"},
	    {"John Doe", 28, "john@test.com"},            
            {"Bob Smith", 35, "jacky@example.com"},
	    {" Swapnil", 37, "joc@example.com."},
	    
        };

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }


        for (int rowNum = 1; rowNum <= data.length; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < headers.length; cellNum++) {
                Cell cell = row.createCell(cellNum);
                Object value = data[rowNum - 1][cellNum];

                if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                }
         
            }
        }

  
        try (FileOutputStream fileOut = new FileOutputStream("sample.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Excel file created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
