package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bouncycastle.operator.AADProcessor;

public class FormattingCellColor {

	public static void main(String[] args) throws IOException {
	
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Stud data");
		
		XSSFRow row = sheet.createRow(1);
		
		//Setting Background Color
		
		XSSFCellStyle style = workbook.createCellStyle();
		
		style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.LESS_DOTS);
        
    	XSSFCell cell = row.createCell(1);
	    cell.setCellValue("Welcome");
	    cell.setCellStyle(style);
        
        //Setting Foregorund Color
	    
	    style = workbook.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.CORAL.getIndex());
	    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    
	    cell = row.createCell(2);
	    cell.setCellValue("Automation");
	    cell.setCellStyle(style);
	    
	    FileOutputStream fos = new FileOutputStream(".\\datafiles\\Styles.xlsx");
	    
	    workbook.write(fos);
	    workbook.close();
	    fos.close();
	    
	    System.out.println("Succesfully style sheet created !!!");
	
	}

	
	
}