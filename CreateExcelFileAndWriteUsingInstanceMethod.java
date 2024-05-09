package Apachi_POI_ReadAndWriteExcelFile;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcelFileAndWriteUsingInstanceMethod {

	public static void main(String[] args) throws IOException {
		
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			XSSFSheet sheet = workbook.createSheet("Emp Details");
			
			Object emp [][] = {
					{1,"Mayuri","Hadapsar"},
					{2,"Shyam","Hadapsar"},
					{3,"Rahul","Kondhwa"}	};
			
			int rows = emp.length;
			int cols = emp[0].length;
			
			System.out.println("Row count : "+rows+ " Col Count : "+cols);

			for(int r=0;r<rows;r++)
			{
				XSSFRow row = sheet.createRow(r);
				for(int c=0;c<cols;c++)
				{
					XSSFCell cell = row.createCell(c);
					Object value = emp[r][c];
					
					if(value instanceof String)
						cell.setCellValue((String)value);
					if(value instanceof Integer)
						cell.setCellValue((Integer)value);
					//if(value.instanceof Double)
						//cell.setCellValue((Double)value);
					
				}
			}
			String filePath = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\Selenium_Session_Practise\\src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\Employee.xlsx";

			FileOutputStream outPutStream = new FileOutputStream(filePath);
			workbook.write(outPutStream);
			
			outPutStream.close();
		}
		System.out.println("Employee.xlsx file is written successfully");
	}

}
