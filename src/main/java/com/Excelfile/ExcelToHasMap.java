package com.Excelfile;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToHasMap {
	
	
	public static void main(String[] arg) throws IOException {
		
	
			FileInputStream fis= new FileInputStream(".\\DataFiles\\HasMapToExcel_Data.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet= workbook.getSheet("Students_Data");
			
			int rows = sheet.getLastRowNum();
			
			//Read Data Excel to HasMap
			
			HashedMap<String, String> data =new HashedMap<String, String>();
			
			for (int r=0; r<=rows; r++) {
				String key=sheet.getRow(r).getCell(0).getStringCellValue();
				String value=sheet.getRow(r).getCell(1).getStringCellValue();
				data.put(key, value);
				
				
			}
			System.out.println("Data    "+data);
			
			
			
			
			
			
			
			
			
		}
	}


