package io.github.nadjannn;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.FileOutputStream;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.function.Consumer;

/**
 * Excel file handling regardless to its type: 97-2004 or latest one.
 */
public abstract class ExcelFileAbstract {

    public static final String DEFAULT_SHEET_NAME = "Sheet1";

    public static final String DEFAULT_DATE_FORMAT = "yyyy-mm-dd";

    /**
     * Processing file name.
     */
    protected final String fileName;

    /**
     * Handling type for processing type: read or write.
     */
    protected final HandlingType handlingType;

    /**
     * Loaded workbook.
     */
    protected final Workbook workbook;

    /**
     * Current sheet instance from excel file. Sheet can be reloaded during file processing.
     */
    protected Sheet sheet;

    /**
     * Current sheet index. Counting starts from 0.
     */
    protected int sheetIndex;

    public ExcelFileAbstract(String fileName, HandlingType handlingType) {
        this.fileName = fileName;
        this.handlingType = handlingType;
        if (handlingType == HandlingType.READ) {
            workbook = loadWorkbook();
            sheet = workbook.getSheetAt(0);
        } else {
            try {
                workbook = createWorkbook();
                sheet = workbook.createSheet(DEFAULT_SHEET_NAME);
            } catch (Exception e) {
                throw new ExcelHandlingException("Could not create new file " + fileName, e);
            }
        }
    }

    public void close() throws ExcelClosingException {
        try {
            if (handlingType == HandlingType.WRITE) {
                save();
            }
            workbook.close();
        } catch (Exception e) {
            throw new ExcelClosingException("Cannot close workbook for file " + fileName, e);
        }
    }

    public int getNumberOfRows() {
        return sheet.getPhysicalNumberOfRows();
    }

    public int getSheetsAmount() {
        return workbook.getNumberOfSheets();
    }

    public void loadSheet(int sheetIndex) {
        sheet = workbook.getSheetAt(sheetIndex);
        this.sheetIndex = sheetIndex;
    }

    public void addAndLoadSheet() {
        int amount = getSheetsAmount();
        workbook.createSheet("Sheet" + (amount + 1));
        loadSheet(amount);
    }

    public <T> Optional<T> getCellValue(int row, int column) {
        return getCell(row, column, false).map(cell -> (T) getCellValue(cell));
    }

    public Optional<Row> getExcelRow(int row) {
        return getExcelRow(row, false);
    }

    public short getLastColumnNumber(int row) {
        return getExcelRow(row).map(Row::getLastCellNum).orElse((short) 0);
    }

    public String getSheetName() {
        return workbook.getSheetName(sheetIndex);
    }

    public void setCellValueString(int row, int column, String value) {
        setDataToCell(row, column, value, (cell) -> {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(value);
        });
    }

    public void setCellValueDouble(int row, int column, Double value) {
        setDataToCell(row, column, value, (cell) -> {
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(value);
        });
    }

    public void setCellValueBoolean(int row, int column, Boolean value) {
        setDataToCell(row, column, value, (cell) -> {
            cell.setCellType(CellType.BOOLEAN);
            cell.setCellValue(value);
        });
    }

    public void setCellValueDate(int row, int column, Date value, String... format) {
        setDataToCell(row, column, value, (cell) -> {
            CellStyle style = workbook.createCellStyle();
            String formatValue = format == null || format.length == 0 ? DEFAULT_DATE_FORMAT : format[0];
            style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(formatValue));
            cell.setCellStyle(style);
            cell.setCellValue(value);
        });
    }

    public void setCellDropDownList(int row, int column, List<String> options) {
        DataValidationHelper dvHelper = createDataValidationHelper();
        DataValidationConstraint dvConstraint = dvHelper.createExplicitListConstraint(ConverterUtil.convertToArrayWithoutNulls(options));
        CellRangeAddressList addressList = new CellRangeAddressList(row, row, column, column);
        DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

    protected Optional<Cell> getCell(int rowNumber, int columnNumber, boolean createIfNotExists) {
        if (rowNumber < 0 || columnNumber < 0) {
            throw new ExcelHandlingException("Can not read cell[" + rowNumber + ", " + columnNumber + "]");
        }
        try {
            Optional<Row> excelRow = getExcelRow(rowNumber, createIfNotExists);
            Optional<Cell> cell = excelRow.map(row -> row.getCell(columnNumber));
            return !cell.isPresent() && createIfNotExists ? excelRow.map(row -> row.createCell(columnNumber)) : cell;
        } catch (Exception e) {
            throw new ExcelHandlingException("Can not read cell[" + rowNumber + ", " + columnNumber + "]", e);
        }
    }

    protected Optional<Row> getExcelRow(int row, boolean createIfNotExists) {
        if (row < 0) {
            throw new ExcelHandlingException("Could not read row " + row);
        }
        try {
            Optional<Row> excelRow = Optional.ofNullable(sheet.getRow(row));
            return !excelRow.isPresent() && createIfNotExists ? Optional.of(sheet.createRow(row)) : excelRow;
        } catch (Exception e) {
            throw new ExcelHandlingException("Could not read row " + row, e);
        }
    }

    protected <T extends Cell> Object getCellValue(T cell) {
        try {
            switch (cell.getCellType()) {
                case BLANK:
                case _NONE:
                case STRING:
                    return getString(cell);
                case NUMERIC:
                    return cell.getNumericCellValue();
                case BOOLEAN:
                    return cell.getBooleanCellValue();
                case ERROR:
                    return cell.getErrorCellValue();
                case FORMULA:
                    return getFormulaValue(cell);
                default:
                    return null;
            }
        } catch (Exception e) {
            throw new ExcelHandlingException("Can not read cell value", e);
        }
    }

    protected Double getFormulaValue(Cell cell) {
        try {
            return Double.valueOf(cell.getNumericCellValue());
        } catch (IllegalStateException e) {
            // Return zero in case of error value in formula cell.
            return 0.0D;
        }
    }

    protected String getString(Cell cell) {
        return Optional.ofNullable(cell.getRichStringCellValue().getString())
                .map(String::trim)
                .orElse(null);
    }

    protected void save() {
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        } catch (Exception e) {
            throw new ExcelHandlingException("Close excel file if it is opened! Could not save file " + fileName, e);
        }
    }

    protected void setDataToCell(int row, int column, Object value, Consumer<CellBase> consumer) {
        if (value != null) {
            getCell(row, column, true).map(c -> (CellBase) c).ifPresent(consumer);
        }
    }

    protected abstract Workbook loadWorkbook();

    protected abstract Workbook createWorkbook();

    protected abstract DataValidationHelper createDataValidationHelper();

}
