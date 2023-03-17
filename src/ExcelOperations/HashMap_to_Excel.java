package ExcelOperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMap_to_Excel {

	public static void main(String[] args) throws IOException {
	
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet =workbook.createSheet();
		
		
		Map<String,String> data = new HashMap<String,String>();
		data.put("Ajit", "101");
		data.put("Maa", "102");
		data.put("Sweta", "103");
		data.put("Bunty", "104");
		data.put("Papa", "105");
		
		
		int rows= 0;
		
		
		for(Map.Entry entry :data.entrySet()) 
		{
			
			XSSFRow row = sheet.createRow(rows++);
			
			row.createCell(0).setCellValue((String)entry.getKey());
			row.createCell(1).setCellValue((String) entry.getValue());
			
			
		}
		
		String filepath = ".\\ExcelFiles\\HashMap_to_Excel.xslx";
		FileOutputStream fos = new FileOutputStream(filepath);
		
			workbook.write(fos);
			
			fos.close();
			
			System.out.println("Excel written sucessfully");
			
	}

}
