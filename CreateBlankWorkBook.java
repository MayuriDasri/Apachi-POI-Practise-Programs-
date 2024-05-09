package Apachi_POI_ReadAndWriteExcelFile;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateBlankWorkBook {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		
		String filePath ="C:\\Users\\lenovo\\eclipse-workspace\\"
				+ "Nov_selenium_work\\Selenium_Session_Practise\\"
				+ "src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\MathsCalculation.xlsx";
		
		XSSFWorkbook workbook = new XSSFWorkbook();
	
		FileOutputStream out = new FileOutputStream(filePath);
		
		workbook.write(out);
		out.close();
		
		
	}

}
