package excelWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWorkbook1 {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
Workbook workbook = new XSSFWorkbook();
        
        // Create a new sheet with the name "sheet1"
        Sheet sheet = (Sheet) workbook.createSheet("sheet1");
        
        // Create column headers
        Row headerRow = ((org.apache.poi.ss.usermodel.Sheet) sheet).createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Age");
        headerRow.createCell(2).setCellValue("Email");

        // Write data to the sheet
        String[][] data = {
            {"John Doe", "30", "john@test.com"},
            {"Jane Doe", "28", "jane@test.com"},
            {"Bob Smith", "35", "bob@example.com"},
            {"Swapnil", "37", "swapnil@example.com"}
        };

        for (int i = 0; i < data.length; i++) {
            Row row = ((org.apache.poi.ss.usermodel.Sheet) sheet).createRow(i + 1);
            for (int j = 0; j < data[i].length; j++) {
                row.createCell(j).setCellValue(data[i][j]);
            }
        }

        // Save the workbook to a file
        try {
            FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\MADHUMATHI E\\Desktop\\ExcelTask13.xlsx");
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Excel file created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


	}


		
