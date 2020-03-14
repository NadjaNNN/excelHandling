package io.github.nadjannn.excel.handling;

import io.github.nadjannn.excel.handling.hssf.HSSFExcelFile;
import io.github.nadjannn.excel.handling.xssf.XSSFExcelFile;
import org.apache.commons.lang3.StringUtils;

/**
 * Factory for file opening when type is defined based on a file extension.
 */
public class ExcelFileFactory {

    private static final String HSSF_EXTENSION = ".xls";

    private static final String XSSF_EXTENSION = ".xlsx";

    /**
     * Returns Excel file for handling. Returns XSSFFile instance for xlsx
     * files. Returns HSSFFile instance for xls files.
     *
     * @param fileName String value of file name.
     * @param handlingType reading or writing type of file handling.
     * @return ExcelFile instance for reading or writing.
     */
    public static ExcelFile openExcelFile(String fileName, HandlingType handlingType) {
        if (StringUtils.isEmpty(fileName) ||  handlingType  == null) {
            throw new ExcelHandlingException("File name and handling types have to be not empty");
        }
        if (fileName.endsWith(HSSF_EXTENSION)) {
            return new HSSFExcelFile(fileName, handlingType);
        } else if (fileName.endsWith(XSSF_EXTENSION)) {
            return new XSSFExcelFile(fileName, handlingType);
        } else {
            throw new ExcelHandlingException("Unsupported file type");
        }
    }

}
