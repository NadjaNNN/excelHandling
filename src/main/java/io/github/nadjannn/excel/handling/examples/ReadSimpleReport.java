package io.github.nadjannn.excel.handling.examples;

import io.github.nadjannn.excel.handling.ExcelFile;
import io.github.nadjannn.excel.handling.ExcelFileFactory;
import io.github.nadjannn.excel.handling.HandlingType;

import java.text.SimpleDateFormat;
import java.util.Date;

public class ReadSimpleReport {

    public static void main(String[] args) {
        readReport("examples/hssfSimpleReport.xls");
        readReport("examples/xssfSimpleReport.xlsx");
    }

    private static void readReport(String fileName) {
        System.out.println("-----------------------------------------------------------------");
        System.out.println("Data from file " + fileName);
        // File will be closed because it extends from java.lang.AutoCloseable.
        try (ExcelFile excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.READ)) {
            // Read all rows on the first sheet.
            for (int row = 1; row < excelFile.getNumberOfRows(); row++) {
                // Integer ID value.
                Integer id = excelFile.getCellValueDouble(row, 0).map(Double::intValue).orElse(null);
                // String text.
                String text = excelFile.getCellValueString(row, 1);
                // Date value.
                Date date = excelFile.getCellValueDate(row, 2).orElse(null);
                // String timestamp converted to Date.
                Date timestamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(excelFile.getCellValueString(row, 3));
                // Double value for number.
                Double number = excelFile.getCellValueDouble(row, 4).orElse(null);
                // String value for the number without extra zero after decimal point.
                String numberString = excelFile.getCellValueString(row, 4);
                // Formatted string value when number representation depends on local computer settings. It is without zero after decimal point.
                String numberStringFormatted = excelFile.getCellValueString(row, 4, true);
                // Boolean value.
                Boolean bool = excelFile.getCellValueBoolean(row, 5).orElse(null);
                // Number with long value.
                Long longNumber = Long.valueOf(excelFile.getCellValueString(row, 6));
                System.out.println("id=" + id + ", text=" + text + ", date=" + date + ", timestamp=" + timestamp
                        + ", number=" + number + ", string number=" + numberString + ", formatted number=" + numberStringFormatted
                        + ", bool=" + bool + ", longNumber=" + longNumber);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
