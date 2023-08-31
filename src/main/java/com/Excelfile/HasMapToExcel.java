package com.Excelfile;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HasMapToExcel {

	public static void main(String[] args) throws IOException {
		
		
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Students_Data");
		
		Map<String, String> data = new HashMap<String, String>();
		data.put("101", "John");
		data.put("102", "Smit");
		data.put("103", "Scott");
		data.put("104", "Mary");
		
		
		int rowno=0;
		
		for (Map.Entry entry : data.entrySet()) {
			
			//System.out.println("entry-->  " +entry);
			
			XSSFRow row= sheet.createRow(rowno++);
			row.createCell(0).setCellValue((String)entry.getKey());
			//System.out.println("Row created for key --->  " + row);
			row.createCell(1).setCellValue((String)entry.getValue());
			//System.out.println("Row created for Value --->  " + row);
		}
		
		FileOutputStream fos = new FileOutputStream(".\\DataFiles\\HasMapToExcel_Data.xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("Excel Writen Succefully");
		
		
		
		

	}

}
