package com.example.writer;

import com.example.annotation.Column;
import com.example.annotation.Header;
import com.example.annotation.Ordered;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.List;

public abstract class WorkbookWriter<T> {
    protected static final int MAX_SIZE = 1048575;
    protected Workbook workbook;
    protected Sheet sheet;

    private String outputFileName;
    private boolean outputHeader;
    private boolean autosize;
    private int rowPos;

    public WorkbookWriter() {
        this.rowPos = 0;
    }

    public void setOutputHeader() {
        this.outputHeader = true;
        this.rowPos = (rowPos == 0) ? 1 : 1;
    }

    public void setAutosizeColumns() {
        this.autosize = true;
    }

    public boolean isOutputHeader() {
        return outputHeader;
    }

    public boolean isAutosize() {
        return autosize;
    }

    public void outputHeader(Class<T> tClass) {
        Row firstRow = sheet.createRow(0);

        int currentFieldIndex = 0;
        for(Field field : tClass.getDeclaredFields()) {
            int fieldIndex = getFieldIndex(tClass, field, currentFieldIndex);
            Cell cell = firstRow.createCell(fieldIndex);
            String cellValue = getFieldName(field);

            cell.setCellValue(cellValue);

            currentFieldIndex++;
        }
    }

    // freeze header row

    public void freezeRows(int colSpan, int rowSpan) {
        sheet.createFreezePane(colSpan, rowSpan);
    }
    // Autosize columns

    public abstract void autosizeColumns();
    // Saving and mapping excel file to specified class

    public abstract Workbook getWorkbook(Class<T> tClass, List<T> list);

    public void setOutputFileName(String fileName) {
        this.outputFileName = fileName;
    }

    protected int getStartPos() {
        return rowPos;
    }

    protected void saveDataToWorkbook(Class<T> tClass, List<T> list) {
        int rowIndex = getStartPos();
        for(T obj : list) {
            Row row = sheet.createRow(rowIndex);

            int fieldId = 0;
            for(Field field : tClass.getDeclaredFields()) {
                field.setAccessible(true);

                int index = getFieldIndex(tClass, field, fieldId);
                Cell cell = row.createCell(index);
                setCellValue(field, cell, obj);

                fieldId++;
            }
            rowIndex++;
        }
    }

    protected int getFieldIndex(Class<T> tClass, Field field, int currentFieldIndex) {
        if(tClass.isAnnotationPresent(Ordered.class)) {
            return currentFieldIndex;
        } else if (field.isAnnotationPresent(Column.class)) {
            return field.getAnnotation(Column.class).index();
        }
        return currentFieldIndex;
    }

    protected String getFieldName(Field field) {
        Class<Header> headerClass = Header.class;

        if(field.isAnnotationPresent(headerClass)) {
            return field.getAnnotation(headerClass).name();
        }
        return field.getName();
    }

    protected String getOutputFileName(Class<T> tClass) {
        if(outputFileName != null && outputFileName.length() > 0) {
            return outputFileName + ".xlsx";
        }
        return tClass.getSimpleName() + ".xlsx";
    }

    protected void setCellValue(Field field, Cell cell, T obj) {
        try {
            Object value = field.get(obj);

            if (value instanceof Integer ) {
                cell.setCellValue(Double.parseDouble(Integer.toString((int) value)));
            } else if (value instanceof Float) {
                cell.setCellValue(Float.parseFloat(Integer.toString((int) value)));
            } else {
                cell.setCellValue((String) value);
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }
}
