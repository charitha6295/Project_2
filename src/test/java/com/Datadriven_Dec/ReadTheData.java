package com.Datadriven_Dec;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadTheData {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\admin\\Documents\\sample1.xlsx");
		
		FileInputStream fi = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fi);
		
		org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheet("Sheet1");
		
		Row row = sheet.getRow(2);
		
		Cell cell = row.getCell(0);
		
		CellType cellType = cell.getCellType();
		
		////Enum enumeration[String,numerics]
		
		if(cellType.equals(CellType.STRING)) {
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
		}else if (cellType.equals(CellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			System.out.println(numericCellValue);
		}
		}
		
		
		
		
}


