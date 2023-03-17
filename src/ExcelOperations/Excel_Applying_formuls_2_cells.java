package ExcelOperations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Applying_formuls_2_cells {

	public static void main(String[] args) throws IOException {
		
		
		String filepath = ".\\ExcelFils\\writ_Formula_2_cell.xslx";
		
		FileInputStream fis = new FileInputStream(filepath);
		

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		XSSFSheet sheet =workbook.getSheet("Sheet1");
		
		sheet.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)");
		
		fis.close();
		
		FileOutputStream fos = new FileOutputStream(filepath);
		workbook.write(fos);
		
		workbook.close();
		fos.close();
		
		System.out.println("Done!!!");
		
		
		
	}

}
