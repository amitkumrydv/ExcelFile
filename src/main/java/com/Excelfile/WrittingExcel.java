package com.Excelfile;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WrittingExcel {

	public static void main(String[] arg) throws IOException {

		// Workbook--->> Sheet--->> Rows---->> Cell
		try {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp_Info");

		Object empData[][] = { { "Empid", "Name", "Job" }, { 101, "Amit", "tester" }, { 103, "David", "Maneger" },
				{ 104, "scot", "AI" }

		};

		// Use for loop
		/*
		  int rows= empData.length; int colms =empData[0].length;
		  
		  System.out.println("row--> " +rows); System.out.println("cell--> "+ colms);
		  
		  for(int r=0; r<rows; r++) { XSSFRow row = sheet.createRow(r); // Zero row Created here
		  
		  for(int c=0; c<colms; c++) {
		  
		  XSSFCell cell = row.createCell(c); // Zero column created here
		  
		  System.out.println(" r  or c --> " + r+"  "+c); 
		  Object value =empData[r][c];
		  
		  if ( value instanceof String) cell.setCellValue((String) value);
		  
		  if ( value instanceof Integer) cell.setCellValue((Integer) value);
		  
		  if ( value instanceof Boolean) cell.setCellValue((Boolean) value);
		  
		  
		  
		  }
		  
		  }
		 */

		// use for... each loop
		int rowCount = 0;

		for (Object emp[] : empData) {
			XSSFRow row = sheet.createRow(rowCount++);
			int colmnsCount = 0;
			for (Object value : emp) {
				XSSFCell cell = row.createCell(colmnsCount++);

				if (value instanceof String)
					cell.setCellValue((String) value);

				if (value instanceof Integer)
					cell.setCellValue((Integer) value);

				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);

			}

		}

		String filePath = ".\\DataFiles\\Employee.xlsx";
		FileOutputStream outStream = new FileOutputStream(filePath);
		
		workbook.write(outStream);
		outStream.close();

		System.out.println("Data inserted successfull in the Excel file");
}
catch(Exception e) {
	e.printStackTrace();
}
	}

}
