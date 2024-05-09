package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CountNoOfInformation {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		
		String path ="C:\\Users\\lenovo\\eclipse-workspace\\"
				+ "Nov_selenium_work\\Selenium_Session_Practise\\"
				+ "src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\CountData.xlsx";
		
		FileInputStream fis = new FileInputStream(path);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		int rowCount = sheet.getLastRowNum();
		int colCount = sheet.getRow(0).getLastCellNum();
		
		System.out.println("Total No of Rows : "+rowCount);
		System.out.println("Total No of Cols :"+colCount);
		
		for(int r =0; r<rowCount; r++)
		{
			XSSFRow row = sheet.getRow(r);
			
			for(int c = 0 ; c<colCount ; c++)
			{
				XSSFCell cell  = row.getCell(c);
				System.out.println(cell.getNumericCellValue());				
			}
		}
		
		
		
		fis.close();
		
	}

}
