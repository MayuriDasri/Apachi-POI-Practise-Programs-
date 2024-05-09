//package Apachi_POI_ReadAndWriteExcelFile;
//
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.IOException;
//
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class ReadAndWriteDataIntoFile {
//
//	public static void main(String[] args) throws IOException {
//		
//		String filePath = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\"
//				+ "Selenium_Session_Practise\\src\\main\\java\\"
//				+ "Apachi_POI_ReadAndWriteExcelFile\\Maths.xlsx";
//		
//		FileInputStream fis = new FileInputStream(filePath);
//		
//		XSSFWorkbook  workbook = new XSSFWorkbook();
//		XSSFSheet sheet = workbook.getSheetAt(0);
//		
//		xxxxxxxxx
//		
//		int rowCount = sheet.getLastRowNum();
//		int colCount = row.getLastCellNum();
//		
//		int exceldata[];
//		for(int r = 0;r<rowCount ; r++)
//		{
//			XSSFRow row = sheet.getRow(r);
//			for(int c =  0;c<colCount;c++)
//			{
//				XSSFCell cell = row.getCell(c);
//				
//			}
//		}
//		
//		
//		
//		 
//		
//
//	}
//
//}
