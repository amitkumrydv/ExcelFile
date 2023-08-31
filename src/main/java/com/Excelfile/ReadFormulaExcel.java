package com.Excelfile;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFormulaExcel {

	public static void main(String[] args) throws IOException {
		
		String filePath= ".\\DataFiles\\Formula_File.xlsx";
		try {
			FileInputStream file = new FileInputStream(filePath);
			XSSFWorkbook workbook= new XSSFWorkbook(file);
			XSSFSheet sheet= workbook.getSheet("Sheet1");
			
			int rows = sheet.getLastRowNum();
			int colms = sheet.getRow(1).getLastCellNum();
			
			//Row
			for(int r=0; r<rows; r++) {
				XSSFRow row = sheet.getRow(r);

				// Cell
				for(int c=0; c<colms; c++) {
					XSSFCell cell = row.getCell(c);
					
					switch(cell.getCellType()){
						case STRING: System.out.print(cell.getStringCellValue()); break;
						case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
						case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
					//	case FORMULA: System.out.println(cell.getCellFormula());break;	
						case FORMULA: System.out.print(cell.getNumericCellValue()); break;
						
					}
					System.out.print("   |    ");
			
				}System.out.println();
			}
	
		} catch (FileNotFoundException e) {
		
			e.printStackTrace();
		}

	}

}
