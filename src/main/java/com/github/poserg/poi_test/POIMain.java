package com.github.poserg.poi_test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Sergey Popov
 *
 */
public class POIMain {

	private static final String DEPARTMENT_REPORT = "Отчет по заявкам в разрезе ведомств";

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		Workbook wb = new HSSFWorkbook();
		
		Sheet sheet = wb.createSheet(DEPARTMENT_REPORT);
		sheet.setFitToPage(true);

		Row titleRow = sheet.createRow(1);
		Cell cell = titleRow.createCell(0);
		cell.setCellValue(DEPARTMENT_REPORT);
		
		// Date
		Row dateRow = sheet.createRow(3);
		Cell dateLabelCell = dateRow.createCell(0);
		dateLabelCell.setCellValue("Дата формирования отчета:");
		Cell dateCell = dateRow.createCell(1);
		Date today = new Date();
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
		dateCell.setCellValue(formatter.format(today));
		
		// Date interval
		Row dateIntervalRow = sheet.createRow(5);
		Cell dateIntervalLabel = dateIntervalRow.createCell(0);
		dateIntervalLabel.setCellValue("Период:");
		Cell dateInterval = dateIntervalRow.createCell(1);
		
		SimpleDateFormat formatter2 = new SimpleDateFormat("dd.MM.yyyy");
		String date1 = formatter2.format(today);
		String date2 = formatter2.format(today);
		dateInterval.setCellValue(date1 + "-" + date2);
		
		createTable(sheet, 8);
		
		String file = "myfile.xls";
		FileOutputStream out;
		try {
			out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}

	private static void createTable(Sheet sheet, int shift) {
		Row tableHeader = sheet.createRow(shift);
		createHeader(tableHeader);
		
		int tableSize = 5;
		
		
		// Total
		Row totalRow = sheet.createRow(sheet.getLastRowNum() + 1);
		totalRow.createCell(0).setCellValue("ИТОГО:");
		
		Row tableFooter = sheet.createRow(sheet.getLastRowNum() + 1);
		createHeader(tableFooter);
	}

	private static void createHeader(Row tableHeader) {
		tableHeader.createCell(0).setCellValue("Ведомство");
		tableHeader.createCell(1).setCellValue("Всего");
	}

}
