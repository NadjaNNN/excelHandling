package com.poi.integration;

import org.junit.Before;

public class XSSFExcelFileReadingTest extends ExcelFileReadingTest {

    @Before
    public void init() {
        final String fileName = getClass().getClassLoader().getResource("xssfFormat.xlsx").getFile();
        excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.READ);
    }

}