package com.example.loader;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;

/**
 * Default interface for loading Excel Workbook from a file
 */
public interface WorkbookLoader {
    Workbook load(String filePath) throws IOException;

    default Workbook load(File file) throws IOException{
        return this.load(file.getAbsolutePath());
    }
}
