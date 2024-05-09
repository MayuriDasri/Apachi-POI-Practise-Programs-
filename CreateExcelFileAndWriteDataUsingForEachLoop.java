package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcelFileAndWriteDataUsingForEachLoop {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Family");
		
		Object family[][] = {
				{"Member Type","Member Name","Member Age"},
				{"Father In Law","Vyankatesh","68"},
				{"Mother In Law","Uma","66"},
				{"Husband","Ghanshyam","38"},
				{"Mother","Mayuri","35"},
				{"Daughter","Kirtana","10"},
				{"son","Kaartik","2"} };
		
		int rows = family.length;
		int cols = family[0].length;
		
		System.out.println("Row Count : "+rows +"Column Count : "+cols);
		
		int rowCount=0;
		for( Object member[]: family) // copying data to row 
		{
			XSSFRow row = sheet.createRow(rowCount++);
			int colCount=0;
			for(Object value : member) // copying row to cell
			{	
				XSSFCell cell = row.createCell(colCount++);
				
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				
			}
		}
		
		String filePath = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\Selenium_Session_Practise\\src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\Family.xlsx";
		FileOutputStream outputStream = new FileOutputStream(filePath);
		workbook.write(outputStream);
		
		System.out.println("Family.xlsx Created Successfully");
	}

}
