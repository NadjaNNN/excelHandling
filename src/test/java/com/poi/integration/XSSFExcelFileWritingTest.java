package com.poi.integration;

import org.junit.Before;

public class XSSFExcelFileWritingTest extends ExcelFileWritingTest {

    @Before
    public void init() {
        fileName = "xssf.xlsx";
    }

}
