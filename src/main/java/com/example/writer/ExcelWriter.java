package com.example.writer;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelWriter<T> {
    private WorkbookWriter<T> workbookWriter;

    public ExcelWriter(WorkbookWriter<T> workbookWriter) {
        this.workbookWriter = workbookWriter;
    }

    public void save(Class<T> tClass, List<T> list) {
        Workbook workbook = workbookWriter.getWorkbook(tClass, list);

        if(workbookWriter.isOutputHeader()) {
            workbookWriter.outputHeader(tClass);
        }

        if(workbookWriter.isAutosize()) {
            workbookWriter.autosizeColumns();
        }

        try (FileOutputStream os = new FileOutputStream(workbookWriter.getOutputFileName(tClass))) {
            workbook.write(os);
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
    }
}
