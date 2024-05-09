package Apachi_POI_ReadAndWriteExcelFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CountNoOfData1 {
	
		
		@SuppressWarnings("resource")
		public static Object [][] FetchData(String path, String sheetname) throws IOException, InterruptedException
		{ 
			File file = new File(path);
			FileInputStream fis =new FileInputStream(file);
			XSSFWorkbook wb =new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheet(sheetname);

			int rowcount = sheet.getLastRowNum()+1;                
			int columncount = sheet.getRow(0).getLastCellNum(); 
			
			System.out.println(rowcount);     
			System.out.println(columncount);  

			Object [][]data=new Object[rowcount][columncount];   
			DataFormatter df =new DataFormatter();
			for(int i=0; i<rowcount; i++)
			{                 
				for(int j=0; j<columncount; j++)
				{
					data[i][j]=df.formatCellValue(sheet.getRow(i).getCell(j)); 
				}
				System.out.println();
			}
			 
			return data;
		}
		public static void readData() throws IOException, InterruptedException
		{
			String path ="C:\\Users\\lenovo\\eclipse-workspace\\Nov_selenium_work\\"
					+ "Selenium_Practise_Programs\\src\\test\\java"
					+ "\\Apachi_POI_ReadAndWriteExcelFile\\LoginDetails.Xlsx";
			
			String sheetname="Sheet1";
			Object[][] testdata = FetchData(path, sheetname);
			
			System.out.println(testdata.length);
			int count=0;
			for(int i=0; i<testdata.length; i++)
			{
				String data = Arrays.toString(testdata[i]);
				if(data == " ")
				{
					continue;
				}
				else
				{
					System.out.println(data);
					count++;
				}
			}
			System.out.println("printed count is:"+count);
		}
		public static void main(String[] args) throws IOException, InterruptedException 
		{
			readData();
		}

	}

