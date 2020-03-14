package io.github.nadjannn.excel.handling.examples;

import io.github.nadjannn.excel.handling.ExcelFile;
import io.github.nadjannn.excel.handling.HandlingType;
import io.github.nadjannn.excel.handling.hssf.HSSFExcelFile;
import io.github.nadjannn.excel.handling.ExcelFileFactory;

public class WritingTwoSheetsHSSFExcelFileExample {

    public static void main(String[] args) {
        // File will be closed because it extends from java.lang.AutoCloseable.
        try (ExcelFile excelFile = ExcelFileFactory.openExcelFile("twoSheetsHssfFormatFile.xls", HandlingType.WRITE)) {
            int row = 0;
            for (int i = 0; i <= 70000; i++) {
                // Add new sheet when first one is over.
                if (row > HSSFExcelFile.MAX_ROW_INDEX_ON_SHEET) {
                    excelFile.addAndLoadSheet();
                    row = 0;
                }
                excelFile.setCellValueString(row, 0, Integer.toString(i));
                row++;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
