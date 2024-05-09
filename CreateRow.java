package Apachi_POI_ReadAndWriteExcelFile;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateRow {

	@SuppressWarnings("resource")
	public static void main(String[] args)throws Exception {
		
		String filePath ="C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\"
				+ "Selenium_Session_Practise\\src\\main\\java\\"
				+ "Apachi_POI_ReadAndWriteExcelFile\\MathsCalculation.xlsx";
		
		FileOutputStream fos = new FileOutputStream(filePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
		
		
		/* It Will only save 30 bcoz we created only one cell i.e 0
		XSSFRow row1 = sheet.createRow(0); 
	    XSSFCell cell = row1.createCell(0); 
		cell.setCellValue(10);
		cell.setCellValue(20);
		cell.setCellValue(30);
		
		XSSFRow row2 = sheet.createRow(2);
		XSSFCell cell2 = row2.createCell(0);
		cell2.setCellValue("True");
		cell2.setCellValue("False"); */
		
		//To OverCome from this problem follow will be solution
		
		XSSFRow row1 = sheet.createRow(0);
		row1.createCell(0).setCellValue(10);
		row1.createCell(1).setCellValue(20);
		row1.createCell(2).setCellValue(30);
		row1.createCell(3).setCellValue(40);
		
		/*XSSFRow row2 = sheet.createRow(2);
		row2.createCell(1).setCellValue("True");
		row2.createCell(3).setCellValue("False");
		row2.createCell(4).setCellValue("Equal"); */
		
		
		//row.createCell(0).setCellValue("Type of Cell");
	    //row.createCell(1).setCellValue("cell value");
		
	
		
		workbook.write(fos);
		
		

	}

}
