package com.github.poserg.poi_test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author Sergey Popov
 *
 */
public class POIMain {

	private static final String TABLE_HEADER = "table_header";
	private static final String TEXT_STYLE = "text_style";
	private static final String SUBTITLE_STYLE = "subtitle_style";
	private static final String TITLE_STYLE = "title_style";
	private static final String DEPARTMENT_REPORT = "Отчет по заявкам в разрезе ведомств";

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		Workbook wb = new HSSFWorkbook();
		
		Sheet sheet = wb.createSheet(DEPARTMENT_REPORT);
		sheet.setFitToPage(true);
		sheet.setDisplayGridlines(false);
		
		Map<String, CellStyle> styles = createStyles(wb);

		Row titleRow = sheet.createRow(1);
		Cell cell = titleRow.createCell(0);
		cell.setCellValue(DEPARTMENT_REPORT);
		cell.setCellStyle(styles.get(TITLE_STYLE));
        titleRow.setHeightInPoints(45);
        sheet.addMergedRegion(CellRangeAddress.valueOf("$A$2:$L$2"));
		
		// Date
		Row dateRow = sheet.createRow(3);
		Cell dateLabelCell = dateRow.createCell(0);
		dateLabelCell.setCellValue("Дата формирования отчета:");
		dateLabelCell.setCellStyle(styles.get(SUBTITLE_STYLE));
		Cell dateCell = dateRow.createCell(1);
		Date today = new Date();
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
		dateCell.setCellValue(formatter.format(today));
		dateCell.setCellStyle(styles.get(TEXT_STYLE));
		
		// Date interval
		Row dateIntervalRow = sheet.createRow(5);
		Cell dateIntervalLabel = dateIntervalRow.createCell(0);
		dateIntervalLabel.setCellValue("Период:");
		Cell dateInterval = dateIntervalRow.createCell(1);
		dateInterval.setCellStyle(styles.get(TABLE_HEADER));
		sheet.setColumnWidth(1, 50*256);
		
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

	private static Map<String, CellStyle> createStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
		
		CellStyle style;
		
		Font titleFont = wb.createFont();
		titleFont.setUnderline(FontUnderline.SINGLE.getByteValue());
		titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		titleFont.setFontHeightInPoints((short) 16);
		style = wb.createCellStyle();
		style.setFont(titleFont);
		styles.put(TITLE_STYLE, style);
		
		Font subtitleFont = wb.createFont();
		subtitleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		subtitleFont.setFontHeightInPoints((short) 11);
		CellStyle subtitleStyle = wb.createCellStyle();
		subtitleStyle.setFont(subtitleFont);
		styles.put(SUBTITLE_STYLE, subtitleStyle);
		
		Font textFont = wb.createFont();
		textFont.setFontHeightInPoints((short) 11);
		CellStyle textStyle = wb.createCellStyle();
		textStyle.setFont(textFont);
		styles.put(TEXT_STYLE, textStyle);
		
		
		CellStyle tableHeaderStyle = wb.createCellStyle();
		tableHeaderStyle.setFont(subtitleFont);
		tableHeaderStyle.setAlignment(CellStyle.ALIGN_CENTER);
		setBorders(tableHeaderStyle);
		setBackgroud(tableHeaderStyle, IndexedColors.GREY_40_PERCENT);
		styles.put(TABLE_HEADER, tableHeaderStyle);
		
		CellStyle subtableTitle = wb.createCellStyle();
		subtableTitle.setFont(subtitleFont);
		setBackgroud(subtableTitle, IndexedColors.GREY_25_PERCENT);
		setBorders(subtableTitle);
		
		CellStyle tableContentOddStyle = wb.createCellStyle();
		tableContentOddStyle.setFont(textFont);
		setBorders(tableContentOddStyle);
		styles.put("tabel_content_odd", tableContentOddStyle);
		
		CellStyle tableContentEvenStyle = wb.createCellStyle();
		tableContentEvenStyle.setFont(textFont);
		setBackgroud(tableContentEvenStyle, IndexedColors.GREY_25_PERCENT);
		setBorders(tableContentEvenStyle);
		styles.put("table_content_even", tableContentEvenStyle);
		
		CellStyle totalTitleCellStyle = wb.createCellStyle();
		totalTitleCellStyle.setFont(subtitleFont);
		totalTitleCellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
		setBorders(totalTitleCellStyle, CellStyle.BORDER_DOUBLE);
		setBackgroud(totalTitleCellStyle, IndexedColors.GREY_40_PERCENT);
		styles.put("total_title", totalTitleCellStyle);
		
		CellStyle totalCellStyle = wb.createCellStyle();
		totalCellStyle.setFont(subtitleFont);
		totalCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		setBorders(totalCellStyle, CellStyle.BORDER_DOUBLE);
		setBackgroud(totalTitleCellStyle, IndexedColors.GREY_40_PERCENT);
		styles.put("total", totalCellStyle);
		
		return styles;
	}

	/**
	 * @param tableHeaderStyle
	 */
	private static void setBackgroud(CellStyle tableHeaderStyle, IndexedColors color) {
		tableHeaderStyle.setFillForegroundColor(color.getIndex());
		tableHeaderStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	}

	/**
	 * @param tableHeaderStyle
	 */
	private static void setBorders(CellStyle tableHeaderStyle) {
		setBorders(tableHeaderStyle, CellStyle.BORDER_MEDIUM);
	}
	
	private static void setBorders(CellStyle tableHeaderStyle, short style) {
		tableHeaderStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
		tableHeaderStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
		tableHeaderStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
		tableHeaderStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
	}
}
