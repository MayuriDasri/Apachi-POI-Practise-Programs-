package Apachi_POI_ReadAndWriteExcelFile;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WorkbookFactoryReadData {

	public static void main(String[] args) throws IOException {
		
		String filePath = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\Selenium_Session_Practise\\src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\LoginDetails.xlsx";
		FileInputStream file = new FileInputStream(filePath);
		
		
		//To read data of  0th Row and 0th column 
		   String s1 =WorkbookFactory.create(file).getSheet("Login").getRow(0).getCell(0).getStringCellValue();
		    System.out.println(s1);
		    
		    
		//To read data of  0th Row and 1st column 
			String filePath1 = "C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\Selenium_Session_Practise\\src\\main\\java\\Apachi_POI_ReadAndWriteExcelFile\\LoginDetails.xlsx";
			FileInputStream file1 = new FileInputStream(filePath1);
		    
	
		   String s2 =WorkbookFactory.create(file1).getSheet("Login").getRow(0).getCell(1).getStringCellValue();
		    System.out.println(s2);
	}

}
