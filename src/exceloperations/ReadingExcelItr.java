package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelItr {

	public static void main(String[] args) throws IOException {
		
String excelFilePath = ".\\datafiles\\countries.xlsx";
		
		//to open a file
		FileInputStream fis = new FileInputStream(excelFilePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
	
		//Iterator(most popular approach)
        Iterator iterator = sheet.iterator();
        
        while(iterator.hasNext())
        {
        	XSSFRow row = (XSSFRow) iterator.next();
        	Iterator cellIterator = row.cellIterator();
        	
        	while(cellIterator.hasNext())
        	{
        	 XSSFCell cell = (XSSFCell) cellIterator.next();

				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue());
				break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());
				break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue());
				break;
				
				}
				System.out.print(" | ");
			}
			
			System.out.println( );
        		
        	}
        	
        
        }
        
	}


