package com.example.writer;

import com.example.exceptions.ListTooBigException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Iterator;
import java.util.List;

/**
 * XSSF implementation of workbook
 * Good for 100000 < elements, allows for autosizing columns
 * @param <T>
 */
public class XSSFWriter<T> extends WorkbookWriter<T> {

    public XSSFWriter() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet();
    }

    @Override
    public void autosizeColumns() {
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet.getPhysicalNumberOfRows() > 0) {
                Row row = sheet.getRow(sheet.getFirstRowNum());
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    int columnIndex = cell.getColumnIndex();
                    sheet.autoSizeColumn(columnIndex);
                }
            }
        }
    }

    @Override
    public Workbook getWorkbook(Class<T> tClass, List<T> list) {
        if(list.size() > MAX_SIZE) {
            throw new ListTooBigException(list.size());
        }

        saveDataToWorkbook(tClass, list);
        return workbook;
    }



}
