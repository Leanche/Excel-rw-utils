package com.example.exceptions;

public class InvalidSheetException extends RuntimeException {
    public InvalidSheetException(String sheetName) {
        super("Invalid sheet name provided: " + sheetName);
    }
}
