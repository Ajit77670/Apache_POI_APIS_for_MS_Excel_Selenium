package ExcelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writing_to_the_Excel {

	public static void main(String[] args) throws IOException {
		
		
// Workbook --> Sheet --> row--> cells
		
		XSSFWorkbook workbook = new XSSFWorkbook();  // create the workbook
		
		XSSFSheet sheet = workbook.createSheet("Emp Info"); // create the sheet
		
		Object empdata[][] = {  {"EmpID","Name","Job"},				// Using Object[][] 2D-array data structure to hold the data.
								{101,"Ajit","Software Engineer"},
								{102,"Sumit","Financial Analyst"},
								{103,"Maa","Life"},
								{104,"Sweta","carying person"}
								
							};
  

// Using for loop writting the data into the excel.
		

// Pre-requiste to define the row and columns:		
	int rows= empdata.length;  		// gives output as 5 row.
	int cols = empdata[0].length;  // gives output as 3 column.
	//empdata[0]esliye liye hain ki to select any one row we can find the columns attach to it. 
		
		
	for(int r=0;r<=rows;r++) // for loop for row
	{
		XSSFRow row=sheet.createRow(r); // created the row
		
		for(int c=0;c<=cols;c++) // for loop for column
		{
			XSSFCell cell=	row.createCell(c);  // created the column/cells
			
			Object value= empdata[r][c];  // writing the row and column  data into the excel.
			
			if(value instanceof String )
				cell.setCellValue((String)value);
			
			if(value instanceof Integer)
				cell.setCellValue((Integer)value);
			
			if(value instanceof Boolean)
				cell.setCellValue((Boolean)value);
		}
		
		
	}
	String filepath=".\\ExcelFiles\\employees.xslx";
	FileOutputStream outputstream = new FileOutputStream(filepath);
	workbook.write(outputstream);
	
	outputstream.close();
	
	System.out.println("employees.xslx file written sucessfully");
	
	}

}
