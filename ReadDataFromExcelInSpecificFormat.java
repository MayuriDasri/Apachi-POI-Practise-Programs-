package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromExcelInSpecificFormat {

	public static void main(String[] args) throws IOException {

		String filePath = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\Selenium_Session_Practise\\src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\LoginDetails.xlsx";
		
		FileInputStream inputStream = new FileInputStream(filePath); 
		System.out.println(inputStream);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream); // open WorkBook
		XSSFSheet sheet2 = workbook.getSheetAt(1); 	//open Sheet
		
		int rowCount = sheet2.getLastRowNum();		// no of rows
		int colCount = sheet2.getRow(1).getLastCellNum();	//no of cols
		
		for(int r=0;r<=rowCount;r++)
		{
			XSSFRow row = sheet2.getRow(r);	//1st Row
			for(int c=0;c<colCount;c++)
			{
				XSSFCell cell = row.getCell(c);
				
				switch(cell.getCellType())
				{
				case STRING :	
								System.out.print(cell.getStringCellValue());
				break;
				
				case NUMERIC : System.out.print(cell.getNumericCellValue());
				break;
				
				default:		System.out.println("No Case Matches ");	
				break;
				
				}
				System.out.print(" | ");
			}
			System.out.println(" ");
		}
		workbook.close();
	}
}
