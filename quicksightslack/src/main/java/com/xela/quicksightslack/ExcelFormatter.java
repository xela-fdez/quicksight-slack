package com.xela.quicksightslack;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFormatter {
	public static String BasicFormatter(String file, int sheetNumber) throws IOException {
		FileInputStream inputStream = new FileInputStream(file);
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		
		for (int i = 0; i < sheet.getNumMergedRegions(); i++)
			sheet.removeMergedRegion(i);

		while (sheet.getFirstRowNum() != 0)
			sheet.shiftRows(0, sheet.getLastRowNum(), -1);
		
		StringBuilder sb = new StringBuilder(file);
		String extension = sb.substring(sb.indexOf(".xls"), sb.length());
		sb.delete(sb.indexOf(".xls"), sb.length());
		sb.append("_formatted"+extension);
		FileOutputStream outputStream = new FileOutputStream(sb.toString());
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
		
		return sb.toString();
	}
	
	public static String BacklogFormatter(String file, int sheetNumber) throws IOException {
		FileInputStream inputStream = new FileInputStream(file);
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		CellStyle highAgeing = workbook.createCellStyle();
		highAgeing.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00;-#,##0.00"));
		highAgeing.setFillForegroundColor(IndexedColors.CORAL.getIndex());
		highAgeing.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		CellStyle mediumAgeing = workbook.createCellStyle();
		mediumAgeing.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00;-#,##0.00"));
		mediumAgeing.setFillForegroundColor(IndexedColors.YELLOW1.getIndex());
		mediumAgeing.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		CellStyle lowAgeing = workbook.createCellStyle();
		lowAgeing.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00;-#,##0.00"));
		lowAgeing.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		lowAgeing.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		List<String> whitelist = new ArrayList<>();
		whitelist.add("ES");
		whitelist.add("MX");
		whitelist.add("marketplace_group_2");
		boolean isWhitelisted = false;
		int mergedCells = sheet.getNumMergedRegions();
		int i;
		for (i = 0; i < mergedCells; i++)
			sheet.removeMergedRegion(i);
		for (i = 0; i < sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				Cell cell = row.getCell(0);
				if (cell == null) {
					sheet.removeRow(row);
				} else {
					if (!cell.getStringCellValue().equals(""))
						isWhitelisted = false;
					Iterator<String> it = whitelist.iterator();
					while (it.hasNext()) {
						if (((String) it.next()).equals(cell.getStringCellValue()))
							isWhitelisted = true;
					}
					if (!isWhitelisted) {
						sheet.removeRow(row);
						sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
						i--;
					}
				}
				if (isWhitelisted) {
					cell = row.getCell(5);
					if (cell != null && cell.getCellType() == CellType.NUMERIC) {
						double value = cell.getNumericCellValue();
						if (row.getCell(2) != null
								&& row.getCell(2).getStringCellValue().equals("SPONSORED_PRODUCTS")) {
							if (value > 24.0D) {
								cell.setCellStyle(highAgeing);
							} else if (value > 12.0D) {
								cell.setCellStyle(mediumAgeing);
							} else {
								cell.setCellStyle(lowAgeing);
							}
						} else if (value > 12.0D) {
							cell.setCellStyle(highAgeing);
						} else if (value > 6.0D) {
							cell.setCellStyle(mediumAgeing);
						} else {
							cell.setCellStyle(lowAgeing);
						}
					}
				}
			}
		}
		while (sheet.getFirstRowNum() != 0)
			sheet.shiftRows(0, sheet.getLastRowNum(), -1);
		for (i = 0; i < 6; i++)
			sheet.autoSizeColumn(i);
		
		StringBuilder sb = new StringBuilder(file);
		String extension = sb.substring(sb.indexOf(".xls"), sb.length());
		sb.delete(sb.indexOf(".xls"), sb.length());
		sb.append("_formatted"+extension);
		FileOutputStream outputStream = new FileOutputStream(sb.toString());
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
		
		return sb.toString();
	}
}
