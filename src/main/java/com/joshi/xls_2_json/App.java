package com.joshi.xls_2_json;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONObject;

/**
 * Convert Excel sheet to most popular JSON format
 * 
 */

public class App {
	public static void main(String[] args) {
		String str = convertExcelToJSON(new File(
				"C:/Users/5013003096/Desktop/sample.xlsx"));
		System.out.println("json str:" + str);
	}

	public static String convertExcelToJSON(File file) {
		JSONObject json = null;
		try {
			FileInputStream inp = new FileInputStream(file);
			Workbook workbook = WorkbookFactory.create(inp);

			// Get the first Sheet.
			Sheet sheet = workbook.getSheetAt(0);

			// Start constructing JSON.
			json = new JSONObject();

			// Iterate through the rows.
			JSONArray rows = new JSONArray();
			for (Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext();) {
				Row row = rowsIT.next();
				JSONObject jRow = new JSONObject();

				// Iterate through the cells.
				JSONArray cells = new JSONArray();
				for (Iterator<Cell> cellsIT = row.cellIterator(); cellsIT
						.hasNext();) {
					Cell cell = cellsIT.next();
					cells.put(cell.getStringCellValue());
				}
				jRow.put("cell", cells);
				rows.put(jRow);
			}

			// Create the JSON.
			json.put("rows", rows);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// Get the JSON text.
		return json.toString();
	}
}
