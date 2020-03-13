package com.poi.integration;

/**
 * Handling exception.
 */
public class ExcelHandlingException extends RuntimeException {

    public ExcelHandlingException(String message) {
        super(message);
    }

    public ExcelHandlingException(String message, Throwable cause) {
        super(message, cause);
    }

}
