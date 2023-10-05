package excelWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Table.Cell;

public class ReadExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
            // Open the Excel file
            FileInputStream fileInputStream = new FileInputStream("C:\\\\Users\\\\MADHUMATHI E\\\\Desktop\\\\ExcelTask13.xlsx");
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            // Get the first sheet (assuming it's the "sheet1" sheet)
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);

            // Iterate through rows and cells to read and print data
            for (Row row : sheet) {
                for (org.apache.poi.ss.usermodel.Cell cell : row) {
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println(); // New line for each row
            }

            // Close the file input stream
            fileInputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    

	}

}
