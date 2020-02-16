package com.example.writer;

import com.example.exceptions.ListTooBigException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import java.util.List;

/**
 * SXSSF implementation of writer.
 * Made for outputting large data sets.
 * Due to Apache limitation it does not allow for autosizing columns.
 * @param <T>
 */
public class SXSSFWriter<T> extends WorkbookWriter<T> {
    private int rowAccessWindow;

    public SXSSFWriter() {
        this.rowAccessWindow = 100;
        this.workbook = new SXSSFWorkbook(rowAccessWindow);
        this.sheet = workbook.createSheet();
    }

    public SXSSFWriter(int rowAccessWindow) {
        this.rowAccessWindow = rowAccessWindow;
        this.workbook = new SXSSFWorkbook(rowAccessWindow);
        this.sheet = workbook.createSheet();
    }

    @Override
    public void autosizeColumns() {
        System.out.println("Autosize for SXSSF not allowed.");
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
