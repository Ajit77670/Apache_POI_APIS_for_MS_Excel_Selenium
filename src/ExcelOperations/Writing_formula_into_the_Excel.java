package ExcelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writing_formula_into_the_Excel {

	public static void main(String[] args) throws IOException {
		
		String filepath = ".\\ExcelFiles\\Creating_row.xslx";
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet =	workbook.createSheet("Numbers"); 
		
		XSSFRow row =sheet.createRow(0);
		
		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		row.createCell(2).setCellValue(30);
		
		row.createCell(3).setCellFormula("A1*B1*C1");
		
		FileOutputStream fos = new FileOutputStream(filepath);
		workbook.write(fos);
		
		fos.close();
		
		System.out.println("Creating_row.xslx is done!!");
		
		
		

	}

}
