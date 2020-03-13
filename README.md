# Excel files handling #

Library was done to simplify Excel files handling based on `org.apache.poi`.
Both file types are supported: "Excel 97-2004 Workbook (.xls)" and "Excel Workbook (.xlsx)".

### Build requirements ###

* Java 8 (1.8.0_181) or higher
* Maven 3

### How to read from Excel file ###

Excel file should be opened for reading. File type will be detected based on a file extension. Provide the file name with extension, for example, `someName.xls` or `someName.xlsx`. 
```
try (ExcelFile excelFile = ExcelFileFactory.openExcelFile("fileName.xlsx", HandlingType.READ)) {
  // Handle excelFile. For example, read some text from first cell.
  String text = excelFile.getCellValueString(0, 0);
}
```
`com.poi.integration.ExcelFile` extends from `java.lang.AutoCloseable`, so, explicit closing is not needed and workbook will be closed when program is finished with excel file.    
See more examples in folder `src/main/java/com/poi/integration/examples`.

### How to write into Excel file ###

Excel file should be opened for writing. File type is detected based on a file extension as for reading. 
```
try (ExcelFile excelFile = ExcelFileFactory.openExcelFile("fileName.xlsx", HandlingType.WRITE)) {
  // Handle excelFile. For example, write text into first cell.
  excelFile.setCellValueString(0, 0, "some text");
}
```
See more examples in folder `src/main/java/com/poi/integration/examples`.

### License ###

This project is licensed under the MIT license. See the [LICENSE](LICENSE) file for more info.

### Author ###

Nadja Nechaeva, email: nnechae@gmail.com
