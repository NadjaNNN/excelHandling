package io.github.nadjannn.examples;

import io.github.nadjannn.ExcelFile;
import io.github.nadjannn.ExcelFileFactory;
import io.github.nadjannn.HandlingType;

import java.util.Arrays;
import java.util.Date;

public class WritingExcelFileExamples {

    public static void main(String[] args) {
        writeExcelFileExample("hssfFormatFile.xls");
        writeExcelFileExample("xssfFormatFile.xlsx");
    }

    private static void writeExcelFileExample(String fileName) {
        // File will be closed because it extends from java.lang.AutoCloseable. File type is defined based on file extension.
        // File is HSSF type if file name extension is .xls and XSSF type for .xlsx.
        try (ExcelFile excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.WRITE)) {
            // Write some text into first row and first column.
            excelFile.setCellValueString(0, 0, "some text");
            // Write Double into first row and second column.
            excelFile.setCellValueDouble(0, 1, 45.5D);
            // Write Boolean into first row and thirds column.
            excelFile.setCellValueBoolean(0, 2, Boolean.TRUE);
            // Write Date with default format into first row and forth column. Default format is yyyy-mm-dd.
            excelFile.setCellValueDate(0, 3, new Date());
            // Write Date with predefined format into first row and fifth column.
            excelFile.setCellValueDate(0, 4, new Date(), "mm/dd/yyyyy");
            // Write drop down list into first row and sixth column. Set text from possible options into cell.
            excelFile.setCellValueString(0, 5, "one");
            excelFile.setCellDropDownList(0, 5, Arrays.asList("one", "two", "three"));
            // Write double with actual long value into first row and seventh column.
            excelFile.setCellValueDouble(0, 6, new Double(1497797039));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
