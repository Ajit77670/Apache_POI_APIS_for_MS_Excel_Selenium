package ExcelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writing_data_to_Excel_Using_for_each_loop {

	public static void main(String[] args) throws IOException {
		
		
// Workbook --> Sheet --> row--> cells
		
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet = workbook.createSheet();
	
	Object Jobdata[][] = { {"SrNo","Company","Package","Location"},
						{1,"ProductBased","Amazon","Bangalore"},
						{2,"ProductBased2","CGI","Bangalore"},
						{3,"ServiceBased","Coginzant","Bangalore"}
			
					};
	
	
	
	int rowcount=0;	
		for( Object data[] :Jobdata) 
		{
			XSSFRow row =sheet.createRow(rowcount++);
			
			int colcount=0;
			for(Object value:data) 
			{
			XSSFCell cell =	row.createCell(colcount++);
			
			if(value instanceof String)
				cell.setCellValue((String)value);
			
			if( value instanceof Integer)
				cell.setCellValue((Integer)value);
			
			if(value instanceof Boolean)
				cell.setCellValue((Boolean)value);
			}
			
		}
	
	// finally writting workbook data to the excel file	
		
		String filepath = ".\\ExcelFiles\\JobHunting.xslx";
		FileOutputStream outputstream = new FileOutputStream(filepath);
		workbook.write(outputstream);
		
		 outputstream.close();
		 
		 System.out.println("JobHunting.xslx file written sucessfully");
		
	}

}
