package com.Excelfile;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ArraylistExcelData {

	public static void main(String[] arg) throws IOException {

		// Workbook--->> Sheet--->> Rows---->> Cell
		try {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp_Info");

		
		ArrayList<Object[]> empData = new ArrayList<Object[]>();
		
		empData.add(new Object[]{ "Empid", "Name", "Job" });
		empData.add(new Object[]{ 101, "Amit", "tester" });
		empData.add(new Object[]{ 103, "David", "Maneger" });
		empData.add(new Object[]{ 104, "scot", "AI" });
		
		
		// use for... each loop
		int rowCount = 0;

		for (Object[] emp : empData) {
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

