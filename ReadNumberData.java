package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadNumberData {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		
		String path = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\"
				+ "Selenium_Practise_Programs\\src\\test\\java\\Apachi_POI_ReadAndWriteExcelFile"
				+ "\\NumberData.xlsx";
		
		FileInputStream fis = new FileInputStream(path);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rowCount = sheet.getLastRowNum()+1;
		int colCount = sheet.getRow(0).getLastCellNum();
		
		System.out.println("Row Count : "+rowCount);
		System.out.println("Column Count : "+ colCount);
		
		int count=0;
		
		for(int r = 0 ; r<rowCount;r++)
		{
			XSSFRow row = sheet.getRow(r);
			for(int c = 0;c<colCount;c++)
			{
				XSSFCell cell = row.getCell(c);
				
				if(cell.getCellType() == CellType.BLANK)
				{
					continue ;
				}
				else
				{
					count++;
					System.out.println(cell.getNumericCellValue());
				}
			}
			System.out.println();
		}
		System.out.println("Total Number of Elements are : "+count);
		
		fis.close();
		

	}

}
