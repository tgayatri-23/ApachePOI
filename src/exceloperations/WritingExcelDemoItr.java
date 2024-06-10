package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelDemoItr {

	public static void main(String[] args) throws IOException {

		// Workbook--Sheet--Rows--Cells

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Details");

		Object empdata[][] = { { "EmpID", "Name", "Job" }, { 101, "Jack", "Engineer" }, { 102, "Danny", "Analyst" },
				{ 103, "Smith", "Tester" }, { 104, "Peter", "Manager" }, { 105, "Berry", "Sales" } };

		// using for each loop
		int rowCount = 0;
		for(Object emp[]:empdata)
		{
			XSSFRow row = sheet.createRow(rowCount++);
			
			int columnCount = 0;
			
			for(Object value:emp)
			{
				XSSFCell cell = row.createCell(columnCount++);
				
				if(value instanceof String)
					cell.setCellValue((String)value);
				
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
		String filePath = ".\\datafiles\\employeedetails.xlsx";
		FileOutputStream fos = new FileOutputStream(filePath);
        workbook.write(fos);
        
        fos.close();
        
        System.out.println("Employeedetails.xls file written successfully...");

	}

}
