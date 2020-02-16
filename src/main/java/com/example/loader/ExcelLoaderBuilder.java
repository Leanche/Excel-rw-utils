package com.example.loader;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;

public class ExcelLoaderBuilder<T> {
    private Workbook workbook;
    private int startingRow;

    public ExcelLoaderBuilder(File file) throws IOException {
        this.workbook = new XSSFWorkbookLoader().load(file);
    }

    public ExcelLoaderBuilder(String filePath) throws IOException {
        this.workbook = new XSSFWorkbookLoader().load(filePath);
    }

    public ExcelLoaderBuilder(Class<T> tClass) throws IOException {
        this.workbook = new XSSFWorkbookLoader().load(tClass.getSimpleName() + ".xlsx");
    }

    public ExcelLoaderBuilder<T> setStartingRow(int index) {
        this.startingRow = index;
        return this;
    }

    public ExcelLoaderBuilder<T> setOmitHeaderRow() {
        return this.setStartingRow(1);
    }

    public ExcelLoader<T> build() {
        return new ExcelLoader<>(this);
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    public int getStartingRow() {
        return startingRow;
    }
}
