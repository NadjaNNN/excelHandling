package io.github.nadjannn;

import static org.junit.Assert.assertTrue;

import io.github.nadjannn.xssf.XSSFExcelFile;
import io.github.nadjannn.hssf.HSSFExcelFile;
import org.junit.Test;

public class ExcelFileFactoryTest {

    @Test(expected = ExcelHandlingException.class)
    public void whenFileNameIsNullShouldThrowException() {
        ExcelFileFactory.openExcelFile(null, HandlingType.READ);
    }

    @Test(expected = ExcelHandlingException.class)
    public void whenFileNameIsEmptyShouldThrowException() {
        ExcelFileFactory.openExcelFile("", HandlingType.READ);
    }

    @Test(expected = ExcelHandlingException.class)
    public void whenHandlingTypeIsNullShouldThrowException() {
        ExcelFileFactory.openExcelFile(getFullPathName("hssfFormat.xls"), null);
    }

    @Test(expected = ExcelHandlingException.class)
    public void whenFileNameIsNotExcelTypeShouldThrowException() {
        ExcelFileFactory.openExcelFile("some", HandlingType.READ);
    }

    @Test
    public void whenFileNameIsXlsxShouldReturnXSSFFile() {
        ExcelFile excelFile = ExcelFileFactory.openExcelFile(getFullPathName("xssfFormat.xlsx"), HandlingType.READ);
        assertTrue(excelFile instanceof XSSFExcelFile);
    }

    @Test
    public void whenFileNameIsXlsShouldReturnHSSFFile() {
        ExcelFile excelFile = ExcelFileFactory.openExcelFile(getFullPathName("hssfFormat.xls"), HandlingType.READ);
        assertTrue(excelFile instanceof HSSFExcelFile);
    }

    private String getFullPathName(String fileName) {
        return getClass().getClassLoader().getResource(fileName).getFile();
    }

}
