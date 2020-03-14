package io.github.nadjannn.xssf;

import io.github.nadjannn.ExcelFile;
import io.github.nadjannn.ExcelFileAbstract;
import io.github.nadjannn.ExcelHandlingException;
import io.github.nadjannn.HandlingType;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

/**
 * Excel file implementation for type "Excel Workbook (.xlsx)".
 */
public class XSSFExcelFile extends ExcelFileAbstract implements ExcelFile {

    public XSSFExcelFile(String fileName, HandlingType handlingType) {
        super(fileName, handlingType);
    }

    protected Workbook loadWorkbook() {
        try {
            return WorkbookFactory.create(new FileInputStream(fileName));
        } catch (Exception e) {
            throw new ExcelHandlingException("Could not open file for reading " + fileName, e);
        }
    }

    protected Workbook createWorkbook() {
        return new XSSFWorkbook();
    }

    protected DataValidationHelper createDataValidationHelper() {
        return new XSSFDataValidationHelper((XSSFSheet) sheet);
    }

}
