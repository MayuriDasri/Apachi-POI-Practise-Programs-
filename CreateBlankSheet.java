package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateBlankSheet {

	public static void main(String[] args) throws Exception {
	
		String filePath ="C:\\Users\\lenovo\\eclipse-workspace\\"
				+ "Nov_selenium_work\\Selenium_Session_Practise\\"
				+ "src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\MathsCalculation.xlsx";

		FileOutputStream fos = new FileOutputStream(filePath);
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();
		workbook.createSheet("First");
		
		workbook.write(fos);
		
		
		
	}

}
