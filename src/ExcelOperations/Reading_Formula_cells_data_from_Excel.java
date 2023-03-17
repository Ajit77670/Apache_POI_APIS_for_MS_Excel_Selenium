package ExcelOperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading_Formula_cells_data_from_Excel {

	public static void main(String[] args) throws IOException {
	
		
	String filepath = ".\\ExcelFils\\Slary.xlsx";
	FileInputStream  inputstream = new FileInputStream(filepath);
	
			XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
			
			XSSFSheet sheet =workbook.getSheet("Sheet1");
	
			
			
		int rows=sheet.getLastRowNum();
		int cols=sheet.getRow(0).getLastCellNum();
		
		for(int r=0;r<rows;r++) 
		{
		XSSFRow row =	sheet.getRow(r);
			
			for(int c=0;c<cols;c++) 
			{
			XSSFCell cells =	row.getCell(c);
			
			switch (cells.getCellType()) 
			{
			
			case STRING : System.out.print(cells.getStringCellValue());break;
			
			case NUMERIC : System.out.print(cells.getNumericCellValue());break;
			
			case BOOLEAN : System.out.print(cells.getBooleanCellValue()); break;
			
			case FORMULA : System.out.print(cells.getNumericCellValue());break;
			
			
			}
			
			System.out.println(" |  ");
		}
			
			System.out.println();
		}
		
		inputstream.close();
	}
	
}


