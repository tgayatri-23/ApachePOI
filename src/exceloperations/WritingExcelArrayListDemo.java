package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelArrayListDemo {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Demo");

		// ArrayList
		ArrayList<Object[]> empdata = new ArrayList<Object[]>();
		empdata.add(new Object[] { "Empid", "Name", "Job" });
		empdata.add(new Object[] { 101, "Sonam", "Actor" });
		empdata.add(new Object[] { 102, "Sachin", "QA" });
		empdata.add(new Object[] { 103, "Madhav", "Doctor" });
		empdata.add(new Object[] { 104, "Saini", "Teacher" });

		// using for each loop
		int rownum = 0;

		for (Object[] emp : empdata) {
			XSSFRow row = sheet.createRow(rownum++);

			int cellnum = 0;
			for (Object value : emp) {

				XSSFCell cell = row.createCell(cellnum++);

				if (value instanceof String)
					cell.setCellValue((String) value);

				if (value instanceof Integer)
					cell.setCellValue((Integer) value);

				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);

			}

		}

		String filePath = ".\\datafiles\\empDemo.xlsx";
		FileOutputStream fos = new FileOutputStream(filePath);
		workbook.write(fos);

		fos.close();

		System.out.println("EmployeeDemo.xls file written successfully...");

	}

}
