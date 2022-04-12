package io.github.nadjannn.excel.handling;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;
import java.util.List;
import java.util.Optional;

/**
 * Excel file with all provided functionality to handle it.
 */
public interface ExcelFile extends AutoCloseable {

    /**
     * Returns number os rows from current loaded sheet.
     *
     * @return Integer value of rows.
     */
    int getNumberOfRows();

    /**
     * Returns columns amount in particular row.
     *
     * @param row Integer value of row counted from zero.
     * @return short value with columns amount in particular row.
     */
    short getLastColumnNumber(int row);

    /**
     * Returns sheets amount from processing file.
     *
     * @return Integers value of sheets amount.
     */
    int getSheetsAmount();

    /**
     * Returns current loaded sheet's name.
     *
     * @return String value of current loaded sheet's name.
     */
    String getSheetName();

    /**
     * Loads sheet with particular index. Sheets counting starts from 0. Loaded sheet is ready for the data processing.
     *
     * @param sheetIndex Integer number for sheet loading.
     */
    void loadSheet(int sheetIndex);

    /**
     * Adds one more sheet and loads it. Loaded sheet is ready for the data processing.
     */
    void addAndLoadSheet();

    /**
     * Returns Optional value with Row instance if such row is present on current sheet.
     *
     * @param row Integer value of row number counted from zero.
     * @return Optional value with Row instance if such row is present on current sheet.
     */
    Optional<Row> getExcelRow(int row);

    /**
     * Reads value from particular cell on current loaded sheet.
     *
     * @param row Integer row value counted from zero.
     * @param column Integer column value counted from zero.
     * @param <T> Cell value type, it can be String for text value, Double for numerical, Boolean for boolean value or Byte with error code.
     * @return Optional with some value from particular cell, which is String for text, Double for numerical types or formulas, Byte for error codes or Boolean.
     */
    <T> Optional<T> getCellValue(int row, int column);

    /**
     * Reads value from cell, returns empty string is cell value is undefined. Cell value is converted to String.
     *
     * @param row    Integer row value counted from zero.
     * @param column Integer column value counted from zero.
     * @return String value from the cell.
     */
    default String getCellValueString(int row, int column) {
        return getCellValueString(row, column, false);
    }

    /**
     * Reads value from cell, returns empty string is cell value is undefined. Cell value is converted to String.
     * Removes extra zeros after dot for numerical values.
     * Applies local computer settings for numerical values representation if parameter format is true.
     *
     * @param row    Integer row value counted from zero.
     * @param column Integer column value counted from zero.
     * @param format boolean value. It is affected on numerical values only, decimal point view will be taken from local computer settings if this parameter is true. Decimal point is always just a dot for numerical values if this parameter is false.
     * @return String value from the cell.
     */
    default String getCellValueString(int row, int column, boolean format) {
        return getCellValue(row, column).map(v -> ConverterUtil.convertToString(v, format)).orElse("");
    }

    /**
     * Returns Optional Double value if cell had a number. Optional is an empty if cell is not numerical.
     *
     * @param row    Integer row value counted from zero.
     * @param column Integer column value counted from zero.
     * @return Optional Double value.
     */
    default Optional<Double> getCellValueDouble(int row, int column) {
        return getCellValue(row, column)
                .map(v -> (v instanceof Double) ? (Double) v : null);
    }

    /**
     * Returns Optional Boolean value if cell had a boolean. Optional is an empty if cell is not boolean.
     *
     * @param row    Integer row value counted from zero.
     * @param column Integer column value counted from zero.
     * @return Optional Boolean value.
     */
    default Optional<Boolean> getCellValueBoolean(int row, int column) {
        return getCellValue(row, column)
                .map(v -> (v instanceof Boolean) ? (Boolean) v : null);
    }

    /**
     * Returns Optional Date value if cell had a number or date, number is converted into Date.
     * Optional is an empty if cell is not numerical or date type.
     *
     * @param row    row Integer row value counted from zero.
     * @param column Integer column value counted from zero.
     * @return Optional Date value.
     */
    default Optional<Date> getCellValueDate(int row, int column) {
        return getCellValue(row, column)
                .map(v -> (v instanceof Double) ? DateUtil.getJavaDate((Double) v) : null);
    }

    /**
     * Sets text value into particular cell if it is not null or empty.
     *
     * @param row    Integer value of row number counted from zero.
     * @param column Integer value of column number counted from zero.
     * @param value  String value.
     */
    void setCellValueString(int row, int column, String value);

    /**
     * Sets numerical value into particular cell if it is not null.
     *
     * @param row    Integer value of row number counted from zero.
     * @param column Integer value of column number counted from zero.
     * @param value  Double numerical value.
     */
    void setCellValueDouble(int row, int column, Double value);

    /**
     * Sets boolean value into particular cell if it is not null.
     *
     * @param row    Integer value of row number counted from zero.
     * @param column Integer value of column number counted from zero.
     * @param value  Boolean value.
     */
    void setCellValueBoolean(int row, int column, Boolean value);

    /**
     * Sets Date into cell.
     *
     * @param row    Integer value of row number counted from zero.
     * @param column Integer value of column number counted from zero.
     * @param value  Date value
     * @param format Optional value for format. Default format is yyyy-mm-dd. Possible date-time formats "m/d/yy", "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yy h:mm".
     */
    void setCellValueDate(int row, int column, Date value, String... format);

    /**
     * Sets drop down list with defined options.
     *
     * @param row     Integer value of row number counted from zero.
     * @param column  Integer value of column number counted from zero.
     * @param options List of String values with options for drop down list.
     */
    void setCellDropDownList(int row, int column, List<String> options);

    /**
     * Closes workbook to release it.
     *
     * @throws ExcelClosingException throws ExcelClosingException in case of error.
     */
    @Override
    void close() throws ExcelClosingException;

    /**
     * Returns file name of current processing file
     *
     * @return String value of processing file
     */
    String getFileName();

    /**
     * Returns handling type of current processing profile. It can be read or write.
     *
     * @return handling type
     */
    HandlingType getHandlingType();

    /**
     * Returns workbook object.
     *
     * @return Workbook object
     */
    Workbook getWorkbook();

    /**
     * Returns current processing sheet
     *
     * @return Sheet with current processing sheet
     */
    Sheet getCurrentSheet();

    /**
     * Returns current processing sheet index, numeration is staring from zero
     *
     * @return int value of current sheet index
     */
    int getCurrentSheetIndex();
}
