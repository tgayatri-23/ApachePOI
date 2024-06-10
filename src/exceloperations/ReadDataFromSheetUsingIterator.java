package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromSheetUsingIterator {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		FileInputStream fis = new FileInputStream(".\\datafiles\\PwdProtected.xlsx");
		String password = "Gayatri@123";

		// XSSFWorkbook workbook = new XSSFWorkbook(fis);

		// use new methor workbookfactory
		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis, password);
		XSSFSheet sheet = workbook.getSheetAt(0);

		// read data from sheet using iterator
		Iterator<Row> iterator = sheet.iterator();
		while (iterator.hasNext()) {
			Row nextrow = iterator.next();

			Iterator<Cell> celliterator = nextrow.cellIterator();

			while (celliterator.hasNext()) {
				Cell cell = celliterator.next();

				switch (cell.getCellType()) {

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
