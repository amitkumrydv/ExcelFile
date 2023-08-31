package com.Excelfile;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PasswordProtectedExcel {
	
	
	public static void main (String arg[]) throws IOException {
		
		
		String FilePath = ".\\DataFiles\\password_protected.xlsx";
		
		
		FileInputStream fs = new FileInputStream(FilePath);
		String password= "123";
		
		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fs,password);
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rows = sheet.getLastRowNum();
		System.out.println("Number of rows--> " + rows);
		int colmn = sheet.getRow(0).getLastCellNum();
		System.out.println("Number of Colmns -------> " + colmn);
		
		Iterator <Row> iterator = sheet.iterator();
		
		while ( iterator.hasNext()) {
			
			XSSFRow row = (XSSFRow) iterator.next();
			
			Iterator<Cell> iteratorcell = row.cellIterator();
			
			while (iteratorcell.hasNext()){
				
				XSSFCell cell = (XSSFCell) iteratorcell.next();
				
				// Switch Case
				
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
				
			}System.out.println();
			
			
			
		}
		
		
		
		
	}

}
