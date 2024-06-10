package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
	
		String excelFilePath = ".\\datafiles\\countries.xlsx";
		
		//to open a file
		FileInputStream fis = new FileInputStream(excelFilePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
	//	XSSFSheet sheet = workbook.getSheet("Sheet1");
	//	instead of sheet name you can use index value also
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//Using For loop
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();
		
		//representing rows
		for(int r=0;r<=rows;r++)
		{
			XSSFRow row = sheet.getRow(r);
			
			
			//representing cols
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell = row.getCell(c);
				
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
