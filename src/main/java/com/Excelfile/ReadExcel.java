package com.Excelfile;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator.*;

import org.apache.commons.math3.util.MultidimensionalCounter.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] arg) throws IOException {

		String excelFilePath = ".\\DataFiles\\countries.xlsx";
		FileInputStream inputStream = new FileInputStream(excelFilePath);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

		// XSSFSheet sheet = workbook.getsheet("List of National Capitals"); // get from
		// sheet name

		XSSFSheet sheet = workbook.getSheetAt(0); // get sheet from the index

		// use for loop

		/*
		 * 
		 * int rows= sheet.getLastRowNum(); 
		 * int cols=sheet.getRow(1).getLastCellNum();
		 * 
		 * for(int r=0; r<=rows; r++) {
		 * 
		 * XSSFRow row=sheet.getRow(r);
		 * 
		 * for( int c=0; c<cols; c++) { 
		 * XSSFCell cell= row.getCell(c);
		 * 
		 * switch(cell.getCellType()) {
		 * 
		 * case STRING: System.out.print(cell.getStringCellValue()); break;
		 *  case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
		 *   case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
		 *    }
		 * System.out.print("  |  ");
		 * 
		 * }
		 * 
		 * System.out.println(); 
		 * 
		 * }
		 * 
		 */

		// Use Iterator
		// Iterator for Row
		java.util.Iterator<Row> iterator = sheet.iterator();

		while (iterator.hasNext()) {

			XSSFRow row = (XSSFRow) iterator.next();
			
		// Iterator for cell
			java.util.Iterator<Cell> cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {
				XSSFCell cell = (XSSFCell) cellIterator.next();

				switch (cell.getCellType()) {

				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				}
				System.out.print("  |  ");

			}
			System.out.println();
			

		}

	}

}
