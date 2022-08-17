package com.HttpStatusCode;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtilities {

	public StringBuilder getColumnData(String ExcelPath, String ExcelName, String sheetName, String columnHeader) {

		String file = ExcelPath + ExcelName;
		StringBuilder content = new StringBuilder("");

		try {
			XSSFWorkbook wb = new XSSFWorkbook(file);
			XSSFSheet sheet = wb.getSheet(sheetName);
			XSSFRow row;
			XSSFCell cell;

			int maxrows; // No of rows
			maxrows = sheet.getPhysicalNumberOfRows();

			int datacol = 0; // No of columns
			int datarow = 0;

			int cellsinrow = 0;
			for (int i = 0; i <= maxrows; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					cellsinrow = row.getPhysicalNumberOfCells();

					int cellCounter = 0;
					while (cellsinrow > 0) {
						cell = row.getCell(cellCounter);
						if (cell != null) {
							cellsinrow--;
							if (cell.getStringCellValue().equals(columnHeader)) {
								datacol = cell.getColumnIndex();
								datarow = i;
								i = maxrows + 1;
								break;
							}
						}
						cellCounter++;
					}

				}
			}

			for (int i = datarow + 1; i <= maxrows + datarow; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					cell = row.getCell(datacol);
					if (cell != null) {
						content.append(cell.getStringCellValue() + "\n");
					}
				}
			}

			wb.close();

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}

		return content;
	}

	public ArrayList<String> getColumnDataAsList(String ExcelPath, String ExcelName, String sheetName,
			String columnHeader) {

		String file = ExcelPath + ExcelName;
		ArrayList<String> content = new ArrayList<String>();

		try {
			XSSFWorkbook wb = new XSSFWorkbook(file);
			XSSFSheet sheet = wb.getSheet(sheetName);
			XSSFRow row;
			XSSFCell cell;

			int maxrows; // No of rows
			maxrows = sheet.getPhysicalNumberOfRows() + 20;

			int datacol = 0; // No of columns
			int datarow = 0;

			int cellsinrow = 0;
			for (int i = 0; i <= maxrows; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					cellsinrow = row.getPhysicalNumberOfCells();

					int cellCounter = 0;
					while (cellsinrow > 0) {
						cell = row.getCell(cellCounter);
						if (cell != null) {
							cellsinrow--;
							if (cell.getStringCellValue().equals(columnHeader)) {
								datacol = cell.getColumnIndex();
								datarow = i;
								i = maxrows + 1;
								break;
							}
						}
						cellCounter++;
					}

				}
			}

			for (int i = datarow + 1; i <= maxrows + datarow; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					cell = row.getCell(datacol);
					if (cell != null) {
						if (cell.getStringCellValue().trim().length() > 0)
							content.add(cell.getStringCellValue());
					}
				}
			}

			wb.close();

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}

		return content;
	}

	public void FillSheet(ArrayList<String> al, String[] columns, String newSheetName, String ExcelPath,
			String ExcelName) throws IOException, InvalidFormatException {
		// Create a Workbook
		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

		// Create a Sheet
		Sheet sheet = workbook.createSheet(newSheetName);

		// Create a Font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Create cells
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Create Other rows and cells with employees data
		int rowNum = 1;
		for (String info : al) {

			String tagName = info.split("<@@@>")[0];
			String tagContent = info.split("<@@@>")[1];
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(tagName);
			row.createCell(1).setCellValue(tagContent);
		}

		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(ExcelPath + ExcelName);
		workbook.write(fileOut);
		fileOut.close();

		// Closing the workbook
		workbook.close();
	}
}