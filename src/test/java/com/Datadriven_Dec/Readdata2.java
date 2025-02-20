package com.Datadriven_Dec;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.ObjectInputStream.GetField;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readdata2 {
	public static void main(String[] args) throws IOException {
File f = new File("C:\\Users\\admin\\Documents\\sample1.xlsx");
		
		FileInputStream fi = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fi);
		
		org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheet("Sheet1");
		
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		
		for (int i = 0; i < physicalNumberOfRows; i++) {
			
			Row row = sheet.getRow(i);
			
			int physicalNumberOfCells = row.getPhysicalNumberOfCells();
			
			for (int j = 0; j < physicalNumberOfCells; j++) {
				Cell cell = row.getCell(j);
				
				CellType cellType = cell.getCellType();
				
				if(cellType.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
				}else if (cellType.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					System.out.println(numericCellValue);
					
					int v= (int)numericCellValue;
					System.out.println(v);
				
			}
			
		}
			
		}
	}

}
