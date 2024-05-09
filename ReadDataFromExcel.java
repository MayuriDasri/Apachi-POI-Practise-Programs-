package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromExcel {

	public static void main(String[] args) throws IOException {
		
		String excelFile = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work"
				+ "\\Selenium_Session_Practise\\src\\main\\java\\"
				+ "Apachi_POI_ReadAndWriteExcelFile\\LoginDetails.xlsx";
		
		FileInputStream inputStream = new FileInputStream(excelFile);
		System.out.println(inputStream);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	
		XSSFSheet sheet = workbook.getSheet("Login");
		
		
		int rowCount = sheet.getLastRowNum();
		int colCount = sheet.getRow(1).getLastCellNum();
		
		for(int r=0;r<=rowCount;r++)
		{
			XSSFRow row = sheet.getRow(r);
			for(int c=0;c<colCount;c++)
			{
				XSSFCell cell = row.getCell(c);	
				
				System.out.print(cell.getStringCellValue());
				System.out.print(" | ");
			}
			System.out.println("");
		}
		
		workbook.close();  
		
	}

}
