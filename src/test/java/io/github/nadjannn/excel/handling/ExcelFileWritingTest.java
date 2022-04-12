package io.github.nadjannn.excel.handling;

import org.apache.poi.ss.usermodel.DateUtil;
import org.junit.After;
import org.junit.Test;

import java.io.File;
import java.util.Date;
import java.util.Optional;
import java.util.function.Consumer;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;

public abstract class ExcelFileWritingTest {

    protected String fileName;

    @After
    public void removeFile() {
        File file = new File(fileName);
        if (file.exists()) {
            new File(fileName).delete();
        }
    }

    @Test
    public void notEmptyStringValueShouldBeSet() throws Exception {
        setStringToFile("text");
        assertEquals("text", getValueFromFile().get());
    }

    @Test
    public void emptyStringValueShouldBeSet() throws Exception {
        setStringToFile("");
        assertEquals("", getValueFromFile().get());
    }

    @Test
    public void nullStringValueShouldNotBeSet() throws Exception {
        setStringToFile(null);
        assertFalse(getValueFromFile().isPresent());
    }

    @Test
    public void notNullNumericalValueShouldBeSet() throws Exception {
        setDoubleToFile(40.7D);
        assertEquals(0, Double.compare(40.7D, (Double) getValueFromFile().get()));
    }

    @Test
    public void nullDoubleValueShouldNotBeSet() throws Exception {
        setDoubleToFile(null);
        assertFalse(getValueFromFile().isPresent());
    }

    @Test
    public void notNullBooleanValueShouldBeSet() throws Exception {
        setBooleanToFile(Boolean.TRUE);
        assertEquals(true, getValueFromFile().get());
    }

    @Test
    public void nullBooleanValueShouldNotBeSet() throws Exception {
        setBooleanToFile(null);
        assertFalse(getValueFromFile().isPresent());
    }

    @Test
    public void nullDateValueShouldNotBeSet() throws Exception {
        setDateToFile(null);
        assertFalse(getValueFromFile().isPresent());
    }

    @Test
    public void notNullDateValueShouldBeSet() throws Exception {
        Date now = new Date();
        setDateToFile(now);
        assertEquals(now, DateUtil.getJavaDate((Double) getValueFromFile().get()));
    }

    @Test
    public void addingANewSheetShouldBeSuccessful() throws Exception {
        applyToFile(excelFile -> excelFile.addAndLoadSheet());
        try (ExcelFile excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.READ)) {
            assertEquals(2, excelFile.getSheetsAmount());
        } catch (Exception e) {
            throw e;
        }
    }

    @Test
    public void getFileNameShouldBeAsInitialised() throws Exception {
        applyToFile(excelFile -> assertEquals(fileName, excelFile.getFileName()));
    }

    @Test
    public void getHandlingTypeShouldBeWrite() throws Exception {
        applyToFile(excelFile -> assertEquals(HandlingType.WRITE, excelFile.getHandlingType()));
    }

    @Test
    public void getWorkbookShouldNotBeNull() throws Exception {
        applyToFile(excelFile -> assertNotNull(excelFile.getWorkbook()));
    }

    @Test
    public void getCurrentSheetShouldNotBeNull() throws Exception {
        applyToFile(excelFile -> assertNotNull(excelFile.getCurrentSheet()));
    }

    @Test
    public void getCurrentSheetIndexShouldBeZero() throws Exception {
        applyToFile(excelFile -> assertEquals(0, excelFile.getCurrentSheetIndex()));
    }

    private void applyToFile(Consumer<ExcelFile> consumer) throws Exception {
        try (ExcelFile excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.WRITE)) {
            consumer.accept(excelFile);
        } catch (Exception e) {
            throw e;
        }
    }

    private void setStringToFile(String text) throws Exception {
        applyToFile(excelFile -> excelFile.setCellValueString(0, 0, text));
    }

    private void setDoubleToFile(Double value) throws Exception {
        applyToFile(excelFile -> excelFile.setCellValueDouble(0, 0, value));
    }

    private void setBooleanToFile(Boolean value) throws Exception {
        applyToFile(excelFile -> excelFile.setCellValueBoolean(0, 0, value));
    }

    private void setDateToFile(Date value) throws Exception {
        applyToFile(excelFile -> excelFile.setCellValueDate(0, 0, value));
    }

    private <T> Optional<T> getValueFromFile() throws Exception {
        try (ExcelFile excelFile = ExcelFileFactory.openExcelFile(fileName, HandlingType.READ)) {
            return excelFile.getCellValue(0, 0);
        } catch (Exception e) {
            throw e;
        }
    }

}
