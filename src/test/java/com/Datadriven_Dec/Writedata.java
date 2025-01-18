package com.Datadriven_Dec;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.poifs.macros.VBAMacroExtractor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writedata {

	public static void main(String[] args) throws IOException {
		
File f = new File("C:\\Users\\admin\\Documents\\sample1.xlsx");
		
		FileInputStream fi = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fi);
		
		//wb.getSheetAt(1).getRow(0).getCell(0).setCellValue("sad");
		
		Sheet sheet = wb.createSheet("Sheet3");
		
		Row row = sheet.createRow(0);
		
		row.createCell(0).setCellValue("hai");
		row.getCell(0).setCellValue("bye");
		
		FileOutputStream fO = new FileOutputStream(f);
		
		wb.write(fO);
		
		System.out.println("done");
	}
}
