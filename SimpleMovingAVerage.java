import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleMovingAVerage {
    public static void main(String[] args) {
        String filePath = "C:\\Project";
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();
            
            int[] indices = {1, 2, 3, 4, 5, 6}; 
            
            for (int index : indices) {
                calculateSMA(sheet, index, 10);
            }
            
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
            
            System.out.println("SMA values updated successfully in the Excel file.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void calculateSMA(Sheet sheet, int colIndex, int period) {
        int lastRow = sheet.getLastRowNum();
        Row headerRow = sheet.getRow(0);
        Cell newHeaderCell = headerRow.createCell(colIndex + 6);
        newHeaderCell.setCellValue(sheet.getRow(0).getCell(colIndex).getStringCellValue() + " SMA(10)");
        
        List<Double> values = new ArrayList<>();
        for (int i = 1; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(colIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    values.add(cell.getNumericCellValue());
                } else {
                    values.add(null);
                }
            }
        }

        for (int i = 1; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            Cell smaCell = row.createCell(colIndex + 6, CellType.NUMERIC);
            if (i >= period) {
                double sum = 0;
                int count = 0;
                for (int j = i - period; j < i; j++) {
                    if (values.get(j) != null) {
                        sum += values.get(j);
                        count++;
                    }
                }
                if (count > 0) {
                    smaCell.setCellValue(sum / count);
                } else {
                    smaCell.setBlank();
                }
            } else {
                smaCell.setBlank();
            }
        }
    }
}
