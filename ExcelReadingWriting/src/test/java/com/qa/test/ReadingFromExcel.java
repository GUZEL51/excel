package com.qa.test;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingFromExcel {

	public static void main(String[] args) throws Exception {
		//capture the path of your excel file
		String path ="C:\\Users\\anach\\Desktop\\SampleExcel1.xlsx";
		//remember to save sample file on desktop
		
		// create an object of file class and pass the path 
		File src = new File(path);
		
		
		//pass this source to the file input system
		FileInputStream fis = new FileInputStream(src);
		
		//create an object of the workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		for (Row row: sheet) {
			for (Cell cell:row) {
				System.out.println(cell.getStringCellValue() + "\t");
			}
			System.out.println();
		
		}
		workbook.close();
		fis.close();
	}

	
	}
