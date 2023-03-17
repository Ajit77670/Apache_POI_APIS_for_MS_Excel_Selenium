package ExcelOperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_to_HashMap {

	public static void main(String[] args) throws IOException {
		
		String filepath = ".\\ExcelFiles\\Excel_to_HashMap.xslx";
		
		FileInputStream fis = new FileInputStream(filepath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		XSSFSheet sheet =workbook.getSheet("Sheet1");
		
		int rows =sheet.getLastRowNum(); // will get the rows by this method.
		
		HashMap<String,String>	data = new HashMap<String,String>();  // created HashMap
		
	//Reading data from Excel to HashMap
		for(int r=0;r<rows;r++) 
		{
		
			String key =sheet.getRow(r).getCell(0).getStringCellValue();
			String value =sheet.getRow(r).getCell(1).getStringCellValue();
		
		
		}
		
	//Reading data from HashMap to console	
		
		
		for(Map.Entry entry:data.entrySet())  // iterate the HasMap on Console
		{

			System.out.println(entry.getKey()+ "   " +entry.getValue());
			
		}

}
}
