package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PasswordProtectedExcel {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream(".\\datafiles\\PwdProtected.xlsx");
		String password = "Gayatri@123";
		
	//	XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		//use new methor workbookfactory
		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis,password);
	    XSSFSheet sheet = workbook.getSheetAt(0);
	    
	    int rows = sheet.getLastRowNum();
	    System.out.println("Rows "  + rows); //count started from 0
	    int cols = sheet.getRow(0).getLastCellNum();
	    System.out.println("Cols " + cols); //count started from 1
	    
	    //read data from sheet using for loop
	    for(int r=0;r<=rows;r++) {
	    	
	    	XSSFRow row = sheet.getRow(r);
	    	
	    	for(int c=0;c<cols;c++){
	    		
	    		XSSFCell cell = row.getCell(c);
	    		
	    		switch(cell.getCellType()) {
	    		
	    		case NUMERIC:
	    			System.out.print(cell.getNumericCellValue());
	    			break;
	    		
	    		case STRING:
	    			System.out.print(cell.getStringCellValue());
	    			break;
	    			
	    		case BOOLEAN:
	    			System.out.print(cell.getBooleanCellValue());
	    			break;
	    			
	    		case FORMULA:
	    			System.out.print(cell.getNumericCellValue());
	    			break;
	    		}
	    		System.out.print(" | ");
	    	}
	    	System.out.println();
	    	
	    	workbook.close();
	    	fis.close();
	    }
	
	
	
	}

}
