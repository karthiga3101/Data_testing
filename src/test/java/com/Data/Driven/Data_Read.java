package com.Data.Driven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data_Read {

	public static void particularData() throws Exception {
		
		File f = new File("D:\\karthika\\Java\\Data.Driven\\data.xlsx"); //xl to file
		FileInputStream fism = new FileInputStream(f);  //get file data
		Workbook wb = new XSSFWorkbook(fism);   //upcasting
		Sheet sheetAt = wb.getSheetAt(0);
		Row row = sheetAt.getRow(1);
		Cell cell = row.getCell(1);
		CellType cellType = cell.getCellType();
		if (cellType.equals(CellType.STRING)) {
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
			
		}
		else if (cellType.equals(CellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			int value = (int) numericCellValue;  //narrow casting
			System.out.println(value);
			
		}
	}
	public static void allData() throws Exception {
		
		File f = new File("D:\\karthika\\Java\\Data.Driven\\data.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		int rowsize = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < rowsize; i++) {
			Row row = sheetAt.getRow(i);
			int cellsize = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cellsize; j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
					
				}
				else if (cellType.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.println(value);
					
				}
			}
				
			}
		
	}
	public static void columnData() throws Exception {
		
		File f = new File("D:\\karthika\\Java\\Data.Driven\\data.xlsx");
		FileInputStream fsm = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fsm);
		Sheet sheetAt = wb.getSheetAt(0);
		int rowsize = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < rowsize; i++) {
			Row row = sheetAt.getRow(i);
			Cell cell = row.getCell(0);
			CellType cellType = cell.getCellType();
			if (cellType.equals(CellType.STRING)) {
				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);
				
			}
			else if (cellType.equals(CellType.NUMERIC)) {
				double numericCellValue = cell.getNumericCellValue();
				int value = (int) numericCellValue;
				System.out.println(value);
				
			}
			
			
		}
		
		}
	
	public static void rowData() throws Exception {
		
		File f = new File("D:\\karthika\\Java\\Data.Driven\\data.xlsx");
		FileInputStream fsm = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fsm);
		Sheet sheetAt = wb.getSheetAt(0);
		Row row = sheetAt.getRow(5);
		int cellsize = row.getPhysicalNumberOfCells();
		for (int i = 0; i < cellsize; i++) {
			Cell cell = row.getCell(i);
			CellType cellType = cell.getCellType();
			if (cellType.equals(CellType.STRING)) {
				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);
				
			}
			else if (cellType.equals(CellType.NUMERIC)) {
				double numericCellValue = cell.getNumericCellValue();
				int value = (int) numericCellValue;
				System.out.println(value);
				
			}
			
			
		}
		

	}
	public static void main(String[] args) throws Throwable {
		System.out.println("****particular data****");
		
		particularData();
		
		System.out.println("****All data****");
		
		allData();
		
		System.out.println("***rowData***");
		
		rowData();
		
		System.out.println("***columndata***");
		
		columnData();
		
	}

}
