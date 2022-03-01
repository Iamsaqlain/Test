package com.example.demo.service;

import java.io.FileOutputStream;

import org.apache.poi.ss.formula.functions.Index;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class TestServiceImpl implements TestService{

	@Override
	public void downloadFile() {
		try {
			FileOutputStream fos = new FileOutputStream("test.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook();
	        XSSFSheet sheet = wb.createSheet("Test");
	        Row row1 = sheet.createRow(0);
	        Cell cell = row1.createCell(0, CellType.STRING);
	        cell.setCellValue("TOP TEN (WITHIN 01-100)");
	    	cell = titleStyle(cell, wb);
	        sheet.addMergedRegion(new CellRangeAddress(0,0,0,3));
	        
	        cell = row1.createCell(4, CellType.STRING);
	        cell.setCellValue("TOP TEN (WITHIN 101-250)");
	        cell= titleStyle(cell, wb);
	        sheet.addMergedRegion(new CellRangeAddress(0,0,4,7));
	        
	        cell = row1.createCell(8, CellType.STRING);
	        cell.setCellValue("TOP TEN (WITHIN 251-500)");
	        cell = titleStyle(cell, wb);
	        sheet.addMergedRegion(new CellRangeAddress(0,0,8,11));
	        
	        Row row2 = sheet.createRow(1);
	        
	        putHeader(row2,cell,wb);
	        
	        Row row3 = sheet.createRow(22);
	        cell = row3.createCell(0, CellType.STRING);
	        cell.setCellValue("BOTTOM TEN (WITHIN 01-100)");
	    	cell = titleStyle(cell, wb);
	        sheet.addMergedRegion(new CellRangeAddress(22,22,0,3));
	        
	        cell = row3.createCell(4, CellType.STRING);
	        cell.setCellValue("BOTTOM TEN (WITHIN 101-250)");
	        cell= titleStyle(cell, wb);
	        sheet.addMergedRegion(new CellRangeAddress(22,22,4,7));
	        
	        cell = row3.createCell(8, CellType.STRING);
	        cell.setCellValue("BOTTOM TEN (WITHIN 251-500)");
	        cell = titleStyle(cell, wb);
	        sheet.addMergedRegion(new CellRangeAddress(22,22,8,11));
	        
	        Row row4 = sheet.createRow(23);
	        
	        putHeader(row4,cell,wb);
	        
	        sheet.setColumnWidth(0, 1500);
	        sheet.setColumnWidth(1, 12000);
	        sheet.setColumnWidth(2, 3000);
	        sheet.setColumnWidth(3, 5000);
	        
	        sheet.setColumnWidth(4, 1500);
	        sheet.setColumnWidth(5, 12000);
	        sheet.setColumnWidth(6, 3000);
	        sheet.setColumnWidth(7, 5000);
	        
	        sheet.setColumnWidth(8, 1500);
	        sheet.setColumnWidth(9, 12000);
	        sheet.setColumnWidth(10, 3000);
	        sheet.setColumnWidth(11, 5000);
	        
	        wb.write(fos);
	 	} catch(Exception ex) {
			ex.printStackTrace();
		}
	}
	
	private void putHeader(Row row,Cell cell, XSSFWorkbook wb) {
		cell = row.createCell(0,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("RANK");
        
        cell = row.createCell(1,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("COMPANY");
        
        cell = row.createCell(2,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("PRICE");
        
        cell = row.createCell(3,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("MARKET CAP (CR)");
        
        cell = row.createCell(4,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("RANK");
        
        cell = row.createCell(5,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("COMPANY");
        
        cell = row.createCell(6,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("PRICE");
        
        cell = row.createCell(7,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("MARKET CAP (CR)");
        
        cell = row.createCell(8,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("RANK");
        
        cell = row.createCell(9,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("COMPANY");
        
        cell = row.createCell(10,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("PRICE");
        
        cell = row.createCell(11,CellType.STRING);
        cell = headereStyle(cell, wb);
        cell.setCellValue("MARKET CAP (CR)");
        
	}

	public Cell titleStyle(Cell cell, XSSFWorkbook wb) {
		CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);  
        Font headerFont = wb.createFont();
        headerFont.setColor(IndexedColors.BLACK.index);
        headerFont.setBold(true);
        style.setFont(headerFont);
        cell.setCellStyle(style);
        return cell;
	}
	
	public Cell headereStyle(Cell cell, XSSFWorkbook wb) {
		CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT); 
        Font headerFont = wb.createFont();
        headerFont.setColor(IndexedColors.WHITE.index);
        style.setFillForegroundColor(IndexedColors.BROWN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(headerFont);
        style.setLeftBorderColor(IndexedColors.GREY_25_PERCENT.index);
        style.setRightBorderColor(IndexedColors.GREY_25_PERCENT.index);
        style.setTopBorderColor(IndexedColors.GREY_25_PERCENT.index);
        style.setBottomBorderColor(IndexedColors.GREY_25_PERCENT.index);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        cell.setCellStyle(style);
        return cell;
	}

}
