package io.github.nadjannn.excel.handling;

import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.assertEquals;

public class HSSFExcelFileReadingTest extends ExcelFileReadingTest {

    @Before
    public void init() {
        final String fileName = getClass().getClassLoader().getResource("hssfFormat.xls").getFile();
        excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.READ);
    }

    @Test
    public void numberOfPhysicalRowsShouldBeTakenFromFile() {
        assertEquals(3, excelFile.getNumberOfRows());
    }

}
