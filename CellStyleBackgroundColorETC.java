package Apachi_POI_ReadAndWriteExcelFile;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;

public class CellStyleBackgroundColorETC {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws Exception {
		
		String filePath ="C:\\Users\\lenovo\\eclipse-workspace\\"
				+ "Nov_selenium_work\\Selenium_Session_Practise\\"
				+ "src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\CellStyle.xlsx";
		
		FileOutputStream fos = new FileOutputStream(filePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("ColorFont");
		
		XSSFRow row = sheet.createRow(0);
		
		XSSFCellStyle style = workbook.createCellStyle();
		
		
		style.setFillBackgroundColor(IndexedColors.BLUE.getIndex());;
		style.setFillPattern(FillPatternType.BIG_SPOTS);
		style.setBorderBottom(null);
		
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("Welcome");
	
		cell.setCellStyle(style);
		
		workbook.write(fos);
		
		


	}

}
