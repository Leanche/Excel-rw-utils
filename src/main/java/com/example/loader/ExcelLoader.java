package com.example.loader;

import com.example.annotation.Column;
import com.example.annotation.Ordered;
import com.example.exceptions.InvalidSheetException;
import com.example.exceptions.StartingRowExceededException;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.*;

public class ExcelLoader<T> {
    private Workbook workbook;
    private int startingRow;

    protected ExcelLoader(ExcelLoaderBuilder<T> builder) {
        this.workbook = builder.getWorkbook();
        this.startingRow = builder.getStartingRow();
    }

    public int getTotalSheetRowSize(int index) {
        return workbook.getSheetAt(index).getPhysicalNumberOfRows();
    }

    public int getTotalSheetRowSize(String sheetName) {
        return workbook.getSheet(sheetName).getPhysicalNumberOfRows();
    }

    public List<T> getList(Class<T> tClass) {
        return this.getListFromSheet(tClass, 0);
    }

    public List<T> getListFromSheet(Class<T> tClass, int sheetId) {
        return this.getListFromSheet(tClass, workbook.getSheetAt(sheetId));
    }

    public List<T> getListFromSheet(Class<T> tClass, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);

        if(sheet == null) {
            throw new InvalidSheetException(sheetName);
        }

        return this.getListFromSheet(tClass, sheet);
    }

    private List<T> getListFromSheet(Class<T> tClass, Sheet sheet) {
        List<T> collection = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.rowIterator();
        int rows = sheet.getPhysicalNumberOfRows();

        if(rows < startingRow) {
            throw new StartingRowExceededException();
        }

        int rowIndex = 0;
        while(rowIterator.hasNext()) {

            if (startingRow > rowIndex) {
                rowIterator.next();
                rowIndex++;
                continue;
            }

            Row row = rowIterator.next();

            T obj = mapRowToClass(tClass, row);

            if(obj != null) {
                collection.add(obj);
            }
        }


        return collection;
    }

    private T mapRowToClass(Class<T> tClass, Row row) {
        try {
            T obj = tClass.newInstance();

            Iterator<Cell> cellIterator = row.cellIterator();

            int cellIndex = 0;
            while(cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                Field[] fields = tClass.getDeclaredFields();

                int fieldIndex = 0;
                for(Field field : fields) {
                    field.setAccessible(true);

                    int index = getFieldIndex(tClass, field, fieldIndex);

                    if(index == cellIndex) {
                        setCellValueToObject(obj, field, cell);
                    }

                    fieldIndex ++;
                }

                cellIndex++;
            }
            return obj;
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        return null;
    }

    private int getFieldIndex(Class<T> tClass, Field field, int fieldIndex) {
        if(tClass.isAnnotationPresent(Ordered.class)) {
            return fieldIndex;
        } else if (field.isAnnotationPresent(Column.class)) {
            return field.getAnnotation(Column.class).index();
        }
        return fieldIndex;
    }

    private void setCellValueToObject(T obj, Field field, Cell cell) throws IllegalAccessException {
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType) {
            case STRING:
                field.set(obj, cell.getStringCellValue());
                break;
            case NUMERIC:
                int value = (int) cell.getNumericCellValue();
                field.set(obj, value);
                break;
        }
    }
}
