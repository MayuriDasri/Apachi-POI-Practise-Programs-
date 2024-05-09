package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCellUsingAvailableExcelSheetValue {

	public static void main(String[] args) throws IOException {
		
		String filePath ="C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work"
				+ "\\Selenium_Session_Practise\\src\\main\\java"
				+ "\\Apachi_POI_ReadAndWriteExcelFile\\MathsCalculation.xlsx";
		
		FileInputStream fis = new FileInputStream(filePath);
		
		try (XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			sheet.getRow(0).getCell(2).setCellFormula("SUM(A1:B1)");
			fis.close();
			
			
			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
		} catch (FormulaParseException | IllegalStateException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("File Updated Successfully");
		
		
		

	}

}
