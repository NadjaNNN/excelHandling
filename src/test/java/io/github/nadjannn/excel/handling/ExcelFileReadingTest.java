package io.github.nadjannn.excel.handling;

import org.apache.poi.ss.usermodel.DateUtil;
import org.junit.After;
import org.junit.Test;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Optional;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

public abstract class ExcelFileReadingTest {

    protected ExcelFile excelFile;

    @After
    public void closeFile() throws ExcelClosingException {
        excelFile.close();
    }

    @Test
    public void numberOfSheetsShouldBeTakenFromFile() {
        assertEquals(2, excelFile.getSheetsAmount());
    }

    @Test
    public void otherSheetShouldBeLoaded() {
        excelFile.loadSheet(1);
    }

    @Test
    public void formulaShouldBeRead() {
        Optional<Object> value = excelFile.getCellValue(3, 0);
        assertTrue(value.isPresent());
        assertEquals(0, Double.compare(20D, ((Double) value.get()).doubleValue()));
    }

    @Test
    public void readingOfFormulaWithErrorShouldReturnZero() {
        Optional<Object> value = excelFile.getCellValue(3, 1);
        assertTrue(value.isPresent());
        assertEquals(0, Double.compare(0D, ((Double) value.get()).doubleValue()));
    }

    @Test
    public void blankTypeCellShouldBeRead() {
        Optional<Object> value = excelFile.getCellValue(3, 2);
        assertTrue(value.isPresent());
        assertEquals("blank", value.get());
    }

    @Test
    public void stringTypeCellShouldBeRead() {
        Optional<Object> value = excelFile.getCellValue(3, 3);
        assertTrue(value.isPresent());
        assertEquals("string", value.get());
    }

    @Test
    public void numericTypeCellShouldBeRead() {
        Optional<Object> value = excelFile.getCellValue(3, 4);
        assertTrue(value.isPresent());
        assertEquals(0, Double.compare(45D, ((Double) value.get()).doubleValue()));
    }

    @Test
    public void booleanTypeCellShouldBeRead() {
        Optional<Object> value = excelFile.getCellValue(3, 5);
        assertTrue(value.isPresent());
        assertEquals(true, value.get());
    }

    @Test
    public void currencyTypeCellShouldBeRead() {
        Optional<Object> value = excelFile.getCellValue(3, 6);
        assertTrue(value.isPresent());
        assertEquals(0, Double.compare(45.5D, ((Double) value.get()).doubleValue()));
    }

    @Test
    public void dateTypeCellShouldBeRead() {
        Optional<Object> value = excelFile.getCellValue(3, 7);
        assertTrue(value.isPresent());
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
        assertEquals("01.02.2020", dateFormat.format(DateUtil.getJavaDate((Double) value.get())));
    }

    @Test
    public void getCellValueStringShouldReturnEmptyStringForEmptyCell() {
        assertEquals("", excelFile.getCellValueString(0, 10));
    }

    @Test
    public void getCellValueStringShouldReturnTextForAllTypes() {
        assertEquals("10", excelFile.getCellValueString(0, 0));
        assertEquals("10.5", excelFile.getCellValueString(0, 1));
        assertEquals("20", excelFile.getCellValueString(3, 0));
        assertEquals("blank", excelFile.getCellValueString(3, 2));
        assertEquals("string", excelFile.getCellValueString(3, 3));
        assertEquals("true", excelFile.getCellValueString(3, 5));
        assertEquals("45.5", excelFile.getCellValueString(3, 6));
        assertEquals("43862", excelFile.getCellValueString(3, 7));
    }

    @Test
    public void getCellValueDoubleShouldReturnEmptyForEmptyCell() {
        assertFalse(excelFile.getCellValueDouble(0, 10).isPresent());
    }

    @Test
    public void getCellValueDoubleShouldReturnEmptyForNonNumericalCells() {
        assertFalse(excelFile.getCellValueDouble(2, 2).isPresent());
        assertFalse(excelFile.getCellValueDouble(2, 3).isPresent());
        assertFalse(excelFile.getCellValueDouble(2, 5).isPresent());
    }

    @Test
    public void getCellValueDoubleShouldReturnValueForNumericalCells() {
        assertEquals(0, Double.compare(10D, excelFile.getCellValueDouble(0, 0).get()));
        assertEquals(0, Double.compare(10.5D, excelFile.getCellValueDouble(0, 1).get()));
        assertEquals(0, Double.compare(20D, excelFile.getCellValueDouble(3, 0).get()));
        assertEquals(0, Double.compare(45.5D, excelFile.getCellValueDouble(3, 6).get()));
        assertEquals(0, Double.compare(43862D, excelFile.getCellValueDouble(3, 7).get()));
    }

    @Test
    public void getCellValueDateShouldReturnEmptyForNonNumericalCells() {
        assertFalse(excelFile.getCellValueDate(3, 2).isPresent());
        assertFalse(excelFile.getCellValueDate(3, 3).isPresent());
        assertFalse(excelFile.getCellValueDate(3, 5).isPresent());
    }

    @Test
    public void getCellValueDateShouldReturnValueForDateCell() {
        Date value = DateUtil.getJavaDate(43862D);
        assertTrue(value.equals(excelFile.getCellValueDate(3, 7).get()));
    }

    @Test
    public void getCellValueBooleanShouldReturnValueForBooleanCell() {
        assertEquals(true, excelFile.getCellValueBoolean(3, 5).get());
    }

    @Test
    public void getCellValueBooleanShouldReturnEmptyForNonBooleanCells() {
        for(int column = 0; column < excelFile.getLastColumnNumber(4); column ++) {
            if (column != 5) {
                assertFalse(excelFile.getCellValueBoolean(4, column).isPresent());
            }
        }
    }

    @Test
    public void lastColumnNumberShouldReturnProperValue() {
        assertEquals(9, excelFile.getLastColumnNumber(3));
    }

    @Test
    public void lastColumnNumberShouldReturnZeroForEmptyRow() {
        assertEquals(0, excelFile.getLastColumnNumber(10));
    }

    @Test
    public void getSheetNameShouldReturnCurrentLoadedSheetName() {
        assertEquals("Sheet1", excelFile.getSheetName());
    }

    @Test
    public void getCellValueStringShouldReadLongValue() {
        assertEquals("10203045689", excelFile.getCellValueString(3,8));
    }

    @Test
    public void getExcelRowShouldReturnRowIfItIsPresent() {
        assertTrue(excelFile.getExcelRow(0).isPresent());
    }

    @Test
    public void getExcelRowShouldNotReturnRowIfItIsNotPresent() {
        assertFalse(excelFile.getExcelRow(100).isPresent());
    }

    @Test
    public void getFileNameShouldNotBeEmpty() {
        assertFalse(excelFile.getFileName().isEmpty());
    }

    @Test
    public void getHandlingTypeShouldBeRead() {
        assertEquals(HandlingType.READ, excelFile.getHandlingType());
    }

    @Test
    public void getWorkbookShouldNotBeNull() {
        assertNotNull(excelFile.getWorkbook());
    }

    @Test
    public void getCurrentSheetShouldNotBeNull() {
        assertNotNull(excelFile.getCurrentSheet());
    }

    @Test
    public void getCurrentSheetIndexShouldBeZero() {
        assertEquals(0, excelFile.getCurrentSheetIndex());
    }

    @Test
    public void lastRowIndexShouldBeTakenFromFile() {
        assertEquals(3, excelFile.getLastRowNum());
    }
}
