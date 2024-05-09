package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCellByAddingDataToExcel {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Addition");
	
		XSSFRow row =sheet.createRow(0);
		
		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		
		row.createCell(2).setCellFormula("A1+B1");
		
		String filePath ="C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\Selenium_Session_Practise\\src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\MathsCalculation.xlsx";
		FileOutputStream outputStream = new FileOutputStream(filePath);
		
		workbook.write(outputStream);
		outputStream.close();
		
		System.out.println("File Updated Successfully");
		
	}

}
