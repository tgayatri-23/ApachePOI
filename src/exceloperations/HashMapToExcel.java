package exceloperations;



import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMapToExcel {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Student Info");
        
		Map<String,String> data = new HashMap<String,String>();
		data.put("101", "Jack");
		data.put("102", "David");
		data.put("103", "Milly");
		data.put("104", "Cherry");
		data.put("105", "Simon");
		
		int rowno = 0;
		
		
		for(Map.Entry entry: data.entrySet())
		{
			XSSFRow row = sheet.createRow(rowno++);
			
			row.createCell(0).setCellValue((String)entry.getKey());
		    row.createCell(1).setCellValue((String) entry.getValue());
			
		} 
		    FileOutputStream fos = new FileOutputStream(".\\datafiles\\Stud.xlsx");
		    
		    workbook.write(fos);
		    fos.close();
		    System.out.println("HashMap to excel written successfully !!!!");
		}
		
		
 }

   

