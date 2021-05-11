package com.Data.Driven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data_Write {
	
	 public static void data1() throws Exception {
		 
		 File f = new File("C:\\Users\\ELCOT\\Desktop\\my selenium\\data1.xlsx");
		 FileInputStream fis = new FileInputStream(f);
		 Workbook wb = new XSSFWorkbook(fis);
		 Sheet cs = wb.createSheet("TimeTable1");
		 Row cr = cs.createRow(0);
		 Cell cc = cr.createCell(0);
		 cc.setCellValue("Day");
		 wb.getSheet("TimeTable1").getRow(0).createCell(1).setCellValue("1st period");
		 wb.getSheet("TimeTable1").getRow(0).createCell(2).setCellValue("2nd period");
		 wb.getSheet("TimeTable1").getRow(0).createCell(3).setCellValue("3rd period");
		 wb.getSheet("TimeTable1").getRow(0).createCell(4).setCellValue("4th period");
		 wb.getSheet("TimeTable1").createRow(1).createCell(0).setCellValue("Monday");
		 wb.getSheet("TimeTable1").getRow(1).createCell(1).setCellValue("Maths");
		 wb.getSheet("TimeTable1").getRow(1).createCell(2).setCellValue("Bio");
		 wb.getSheet("TimeTable1").getRow(1).createCell(3).setCellValue("tamil");
		 wb.getSheet("TimeTable1").getRow(1).createCell(4).setCellValue("English");
		 wb.getSheet("TimeTable1").createRow(2).createCell(0).setCellValue("Tuesday");
		 wb.getSheet("TimeTable1").getRow(2).createCell(1).setCellValue("Bio");
		 wb.getSheet("TimeTable1").getRow(2).createCell(2).setCellValue("Maths");
		 wb.getSheet("TimeTable1").getRow(2).createCell(3).setCellValue("PET");
		 wb.getSheet("TimeTable1").getRow(2).createCell(4).setCellValue("Physics");
		 wb.getSheet("TimeTable1").createRow(3).createCell(0).setCellValue("Wednesday");
		 wb.getSheet("TimeTable1").getRow(3).createCell(1).setCellValue("Social");
		 wb.getSheet("TimeTable1").getRow(3).createCell(2).setCellValue("Maths");
		 wb.getSheet("TimeTable1").getRow(3).createCell(3).setCellValue("Chemistry");
		 wb.getSheet("TimeTable1").getRow(3).createCell(4).setCellValue("Zoo");
		
		 
		 FileOutputStream fos = new FileOutputStream(f);
		 wb.write(fos);
		 wb.close();
		 System.out.println("Process Completed");
		 
		 }
	 public static void main(String[] args) throws Throwable {
		 data1();
		
	}
	 

}
