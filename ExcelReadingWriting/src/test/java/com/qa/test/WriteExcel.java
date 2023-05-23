package com.qa.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws Exception {
	//create a workbook 
		XSSFWorkbook wb= new XSSFWorkbook();
		//create sheet on workbook 
		XSSFSheet sheet1 = wb.createSheet("SheetOne");

		
		for (int rows = 0; rows <10; rows++) {
			Row row =sheet1.createRow(rows);
			for (int cols= 0; cols < 10; cols++) {
				Cell cell =row.createCell(cols);
				
				cell.setCellValue((int)(Math.random()*100));
				
	
		}
		//create file stream
		File dstn = new File("C:\\Users\\anach\\Desktop\\MySampleExcel1.xlsx");
		//chain output stream to path 
		FileOutputStream fos = new FileOutputStream(dstn);
		//write workbook to output stream
		wb.write(fos);
		
		System.out.println("Excel file is written !!");
		
		}
	}
}