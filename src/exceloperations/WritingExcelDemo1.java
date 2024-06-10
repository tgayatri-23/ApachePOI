package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelDemo1 {

	public static void main(String[] args) throws IOException {

		// Workbook--Sheet--Rows--Cells

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");

		Object empdata[][] = { { "EmpID", "Name", "Job" }, { 101, "Jack", "Engineer" }, { 102, "Danny", "Analyst" },
				{ 103, "Smith", "Tester" }, { 104, "Peter", "Manager" }, { 105, "Berry", "Sales" } };
		// Using for loop
		int rows = empdata.length;
		int cols = empdata[0].length;

		System.out.println("Rows: " + rows);
		System.out.println("Cols: " + cols);

		for (int r = 0; r < rows; r++) {
			XSSFRow row = sheet.createRow(r);

			for (int c = 0; c < cols; c++) {
				XSSFCell cell = row.createCell(c);
				Object value = empdata[r][c];

				if (value instanceof String)
					cell.setCellValue((String) (value));

				if (value instanceof Integer)
					cell.setCellValue((Integer) (value));

				if (value instanceof Boolean)
					cell.setCellValue((Boolean) (value));
			}
		}
		String filePath = ".\\datafiles\\employee.xlsx";
		FileOutputStream fos = new FileOutputStream(filePath);
        workbook.write(fos);
        
        fos.close();
        
        System.out.println("Employee.xls file written successfully...");
	}
}
