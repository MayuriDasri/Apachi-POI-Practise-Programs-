package Apachi_POI_ReadAndWriteExcelFile;


import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcelFileAndWriteDataUsingArrayList {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Kids Details");
		
		ArrayList<Object[]> kidsDetails = new ArrayList<Object[]>();

		kidsDetails.add(new Object[] {1,"Kirtana","Fair"});
		kidsDetails.add(new Object[] {2,"Kaartik","Fair"});
		kidsDetails.add(new Object[] {3,"Tanmay","Black"});
		kidsDetails.add(new Object[] {4,"Sai","Black"});
		
		int rownum = 0;
		
		for(Object [] kids : kidsDetails) 
		{
			XSSFRow row = sheet.createRow(rownum++);
			int cellnum = 0;
			for(Object value : kids)
			{
				XSSFCell cell = row.createCell(cellnum++);
				
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
			}
		}
		
		String filePath = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\"
				+ "Selenium_Session_Practise\\src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile"
				+ "\\KidInformation.xlsx";
		
		FileOutputStream output = new FileOutputStream(filePath);
		
		workbook.write(output);
		System.out.println("Kids Information .xlsx file Created and data written Successfuly");
		
	}

}

