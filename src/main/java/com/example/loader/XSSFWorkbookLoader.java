package com.example.loader;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

/**
 * XSSF implementation of loading workbook.
 */
public class XSSFWorkbookLoader implements WorkbookLoader {
    @Override
    public Workbook load(String filePath) throws IOException {
        return new XSSFWorkbook(filePath);
    }
}
