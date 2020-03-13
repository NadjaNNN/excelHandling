package com.poi.integration;

import org.junit.Before;

public class HSSFExcelFileReadingTest extends ExcelFileReadingTest {

    @Before
    public void init() {
        final String fileName = getClass().getClassLoader().getResource("hssfFormat.xls").getFile();
        excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.READ);
    }

}
