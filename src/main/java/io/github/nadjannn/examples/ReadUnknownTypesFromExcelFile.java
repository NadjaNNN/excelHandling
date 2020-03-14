package io.github.nadjannn.examples;

import io.github.nadjannn.ExcelFile;
import io.github.nadjannn.ExcelFileFactory;
import io.github.nadjannn.HandlingType;

public class ReadUnknownTypesFromExcelFile {

    public static void main(String[] args) {
        readUnknownTypes("examples/hssfAllTypesFile.xls");
        readUnknownTypes("examples/xssfAllTypesFile.xlsx");
    }

    private static void readUnknownTypes(String fileName) {
        // File will be closed because it extends from java.lang.AutoCloseable. File type is defined based on file extension.
        // File is HSSF type if file name extension is .xls and XSSF type for .xlsx.
        System.out.println("-----------------------------------------------------------------");
        System.out.println("Data from file " + fileName);
        try (ExcelFile excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.READ)) {
            int columnsAmount = excelFile.getLastColumnNumber(0);
            // Read all columns on first and second rows from first sheet.
            for (int column = 0; column < columnsAmount; column++) {
                // Read title from first row.
                String title = (String) excelFile.getCellValue(0, column).get();
                // Read value from second row.
                Object value = excelFile.getCellValue(1, column).orElse(null);
                System.out.println(title + "=" + value);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
