package ExcelOperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class Reading_from_the_Excel {

	public static void main(String[] args) throws IOException  {
	
		
		String filepath =".\\ExcelFiles\\Upload Format.xlsx";
		
		FileInputStream fis = new FileInputStream(filepath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis); // Workbook decalaration.
		
		XSSFSheet sheet = workbook.getSheet("Sheet1"); //Sheet decalaration.

///////////////////////////// Reading Data in exele sheet using for Loop:////////////////////////////////////////////
		
		
	//Pre-requiste to find the row and colum :	
	int rows =	sheet.getLastRowNum();      		// getLastRowNum() gives the total no of row in the excel sheet.
	int cols =	sheet.getRow(1).getLastCellNum();	// getRow().getLastCellNum() gives the total no of column.
	
	/*
	for(int r=0; r<=rows; r++)       // Representing the rows
	{
		XSSFRow row = sheet.getRow(r);  // selecting the rows
		
		for(int c=0;c<cols;c++)		// Representing the column (column contain cells of the excel sheet)
		{
			
			XSSFCell cell =row.getCell(c);  // with the row we can read the cells data remember this.
			
			switch (cell.getCellType())
			{
				
			case STRING : System.out.print(cell.getStringCellValue()); break;
			
			case NUMERIC : System.out.print(cell.getNumericCellValue());break;
			
			case BOOLEAN : System.out.print(cell.getBooleanCellValue());break;
			
			}
			
			System.out.print(" | ");
		}
		
		System.out.println();
		
		
		fis.close();
	}
	*/
/////////////////////////////////////////////Reading Data in exele sheet using Iterator:///////////////////////////////////////////////
	
	
	
	Iterator iterator =sheet.iterator();  // iterating for row
	
	while(iterator.hasNext())  // while for row
	{
			
		XSSFRow row =(XSSFRow)iterator.next();  // selecting row
		
		Iterator cellIterator =row.cellIterator();   // iterating for column
		
		while(cellIterator.hasNext())   // while for column
		{
		XSSFCell cell=(XSSFCell) cellIterator.next();  //  // selecting cell info w.r.t column
		
		switch (cell.getCellType())
		{
			
		case STRING : System.out.print(cell.getStringCellValue()); break;
		case NUMERIC : System.out.print(cell.getNumericCellValue());break;
		case BOOLEAN : System.out.print(cell.getBooleanCellValue());break;
		
		}
		
		
		System.out.print(" | ");
	}
	
	System.out.println();
			
		}
	}
	
	
	}


