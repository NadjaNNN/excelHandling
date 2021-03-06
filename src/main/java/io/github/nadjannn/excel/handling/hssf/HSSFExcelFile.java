package io.github.nadjannn.excel.handling.hssf;

import io.github.nadjannn.excel.handling.ExcelFile;
import io.github.nadjannn.excel.handling.ExcelFileAbstract;
import io.github.nadjannn.excel.handling.ExcelHandlingException;
import io.github.nadjannn.excel.handling.HandlingType;
import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;

/**
 * Excel file implementation for type "Excel 97-2004 Workbook (.xls)".
 */
public class HSSFExcelFile extends ExcelFileAbstract implements ExcelFile {

    public static final int MAX_ROW_INDEX_ON_SHEET = 65535;

    public HSSFExcelFile(String fileName, HandlingType handlingType) {
        super(fileName, handlingType);
    }

    protected Workbook loadWorkbook() {
        try {
            POIFSFileSystem poiFileSystem = new POIFSFileSystem(new FileInputStream(fileName));
            return new HSSFWorkbook(poiFileSystem);
        } catch (Exception e) {
            throw new ExcelHandlingException("Could not open file for reading " + fileName, e);
        }
    }

    protected Workbook createWorkbook() {
        return new HSSFWorkbook();
    }

    protected DataValidationHelper createDataValidationHelper() {
        return new HSSFDataValidationHelper((HSSFSheet) sheet);
    }

}
