package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcelUsingIterator {
	
	@SuppressWarnings("resource")
	public static void main (String args[]) throws IOException
	{
		String filePath = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\Selenium_Practise_Programs"
				+ "\\src\\test\\java\\Apachi_POI_ReadAndWriteExcelFile\\LoginDetails.Xlsx";
		FileInputStream inputStream = new FileInputStream(filePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rowCount = sheet.getLastRowNum()+1;
		int colCount = sheet.getRow(0).getLastCellNum();
		System.out.println("Row Count "+rowCount);
		System.out.println("Column Count "+colCount);
		
		Iterator<Row> iterator = sheet.iterator();
		
		while(iterator.hasNext())
		{
			XSSFRow row = (XSSFRow) iterator.next();
		
			Iterator<Cell> cellIterator = row.cellIterator();
		
			while(cellIterator.hasNext())
			{
				XSSFCell cell = (XSSFCell) cellIterator.next();
				
				switch(cell.getCellType())
				{
				
				case STRING : System.out.print(cell.getStringCellValue());
				break;
	
				case NUMERIC : System.out.print(cell.getNumericCellValue());
				break;
	
				default:		System.out.println("No Case Matches ");	
				break;
				
				}
				System.out.print(" | ");
			}
			System.out.println(" ");
		}
	}

	
	}

