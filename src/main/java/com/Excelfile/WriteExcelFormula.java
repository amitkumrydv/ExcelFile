package com.Excelfile;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFormula {

	public static void main(String[] args) throws IOException {
		
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("formulaSheet");
		
		XSSFRow rows = sheet.createRow(0);
		rows.createCell(0).setCellValue(10);
		rows.createCell(1).setCellValue(87);
		rows.createCell(2).setCellValue(90);
		
		rows.createCell(3).setCellFormula("A1*B1*C1");
		
		String filePath=".\\DataFiles\\Write_Formula_File.xlsx";
		try {
			FileOutputStream outputStyrream = new FileOutputStream(filePath);
			workbook.write(outputStyrream);
			outputStyrream.close();
			
		} catch (FileNotFoundException e) {
		
			e.printStackTrace();
		}
		
		System.out.println("successfull ");

	}

}
