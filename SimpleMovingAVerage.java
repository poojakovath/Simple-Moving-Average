package com.example.SimpleMovingAverage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleMovingAverage {
	public static void main(String[] args) {
		// Path to the Excel file
		String filePath = "C:\\Project\\SimpleMovingAverage\\RELIANCE-_1_.xlsx";

		try {
			File file = new File(filePath);
			if (!file.exists()) {
				System.out.println("File not found: " + filePath);
				return;
			}

			FileInputStream inputStream = new FileInputStream(file);
			Workbook workbook = new XSSFWorkbook(inputStream);
			Sheet sheet = workbook.getSheetAt(0);

			// Lists to store the prices and volumes
			List<Double> openPrices = new ArrayList<Double>();
			List<Double> highPrices = new ArrayList<Double>();
			List<Double> lowPrices = new ArrayList<Double>();
			List<Double> closePrices = new ArrayList<Double>();
			List<Double> adjClosePrices = new ArrayList<Double>();
			List<Double> volume = new ArrayList<Double>();

			// Read data from the sheet
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				if (row == null)
					continue;

				openPrices.add(getValidCellValue(row.getCell(1)));
				highPrices.add(getValidCellValue(row.getCell(2)));
				lowPrices.add(getValidCellValue(row.getCell(3)));
				closePrices.add(getValidCellValue(row.getCell(4)));
				adjClosePrices.add(getValidCellValue(row.getCell(5)));
				volume.add(getValidCellValue(row.getCell(6)));
			}

			int period = 10;
			List<Double> openSMA = calculateSMA(openPrices, period);
			List<Double> highSMA = calculateSMA(highPrices, period);
			List<Double> lowSMA = calculateSMA(lowPrices, period);
			List<Double> closeSMA = calculateSMA(closePrices, period);
			List<Double> adjCloseSMA = calculateSMA(adjClosePrices, period);
			List<Double> volumeSMA = calculateSMA(volume, period);

			// Write SMA values back to the sheet
			for (int i = 0; i < openPrices.size(); i++) {
				Row row = sheet.getRow(i + 1);
				if (row == null)
					row = sheet.createRow(i + 1);

				setCellValue(row, 7, openSMA.get(i));
				setCellValue(row, 8, highSMA.get(i));
				setCellValue(row, 9, lowSMA.get(i));
				setCellValue(row, 10, closeSMA.get(i));
				setCellValue(row, 11, adjCloseSMA.get(i));
				setCellValue(row, 12, volumeSMA.get(i));
			}

			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();

			System.out.println("SMA values successfully updated in the Excel sheet.");
		} catch (Exception e) {
			System.out.println("An error occurred.");
			e.printStackTrace();
		}
	}

	// Get the valid cell value, if present
	private static double getValidCellValue(Cell cell) {
		if (cell == null)
			return Double.NaN;
		switch (cell.getCellType()) {
		case NUMERIC:
			return cell.getNumericCellValue();
		case STRING:
			try {
				return Double.parseDouble(cell.getStringCellValue().trim());
			} catch (NumberFormatException e) {
				return Double.NaN;
			}
		case FORMULA:
			try {
				return cell.getNumericCellValue();
			} catch (IllegalStateException e) {
				return Double.NaN;
			}
		default:
			return Double.NaN;
		}
	}

	// Calculate the SMA values
	private static List<Double> calculateSMA(List<Double> prices, int period) {
		List<Double> sma = new ArrayList<Double>();
		for (int i = 0; i < prices.size(); i++) {
			if (i < period - 1) {
				sma.add(Double.NaN);
			} else {
				double sum = 0;
				int validCount = 0;
				for (int j = 0; j < period; j++) {
					double price = prices.get(i - j);
					if (!Double.isNaN(price)) {
						sum += price;
						validCount++;
					}
				}
				sma.add(validCount > 0 ? sum / validCount : Double.NaN);
			}
		}
		return sma;
	}

	// Set the cell value in the sheet
	private static void setCellValue(Row row, int cellIndex, Double newValue) {
		Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
		if (Double.isNaN(newValue)) {
			cell.setBlank();
		} else {
			cell.setCellValue(newValue);
		}
	}
}
