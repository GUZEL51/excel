package com.qa.test;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

	public class Read {


	public static void main(String[] args) throws EncryptedDocumentException, IOException {
	
		String filePath = "C:\\Users\\anach\\Desktop\\SampleExcel2.xlsx";
		//getting the file path through a file
		File src =new File(filePath);
		//create an instance of workbook class which represents
		//the excel file. this is done by passing a file object
		Workbook workbook = WorkbookFactory.create(src);
		
		//from workbook object get a sheet which you want to read
		Sheet sheet = workbook.getSheetAt(0);
		
		//iterate through rows and columns of the sheet
		//using nested loops
		

		for (Row row: sheet) {
		 	for (Cell cell: row) {
		 		switch (cell.getCellType()) {
		 		case STRING:
		 			System.out.println(cell.getStringCellValue()+"\t");
		 			break;
		 		case NUMERIC:
		 			System.out.println(cell.getStringCellValue()+"\t");
		 		case BOOLEAN:
		 		System.out.println(cell.getStringCellValue()+"\t");
		 		break;
		 		}
		 		
		 	}
		 		
		 	//System.out.println(cell.getStringCellValue()+ "\t");
		 		System.out.println();
		 	}
		}
	}	


