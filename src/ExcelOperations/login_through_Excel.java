package ExcelOperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class login_through_Excel {

	public static void main(String[] args) throws IOException {
		
		
		//WebDriverManager.chrome.setup();
		//WebDriver driver = new ChromeDriver();
		//WebElement username = driver.findElementBy(By.id("username"));
		//WebElement pwd = driver.findElementBy(By.id("password"));
		
		
		String path ="C:\\Users\\Ajith Kumar\\eclipse-workspace\\Apache_POI_APIS_for_MS_Excel_Selenium\\ExcelFiles\\login.xlsx";
		
		FileInputStream fis = new FileInputStream(path);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		int rows =sheet.getLastRowNum();  // Total no. of row
		int cols =sheet.getRow(0).getLastCellNum(); // total no. of column.
		
		for(int r =1; r<rows;r++) {
			
		XSSFRow row =	sheet.getRow(rows);
		
		
		for(int c=0; c<cols;c++) {
			
			XSSFCell cells =row.getCell(cols);
			
			switch (cells.getCellType()) {
			
			
			case STRING : System.out.println(cells.getStringCellValue());
			
			case NUMERIC : System.out.println(cells.getBooleanCellValue());
			
			case BOOLEAN : System.out.println(cells.getBooleanCellValue());
			
			case FORMULA : System.out.println(cells.getNumericCellValue());
			
			//username.sendkeys();
			//pwd.sendkeys();
			
			}
		}
		
			
		}
		// Note : This is not the correct programe.

	}

}
