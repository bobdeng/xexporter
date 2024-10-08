package cn.beagile.xexporter;

import lombok.SneakyThrows;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;


class ExcelExportTest {
    @SneakyThrows
    @Test
    void 只有一个单元格() {
        ExcelExport excelExport = new ExcelExport();
        ExcelRow row = new ExcelRow();
        row.setHeight(50);
        excelExport.addRow(row);
        ExcelCell cell = new ExcelCell("这是一个单元格,,很长很长很长", 30, 14);
        row.addCell(cell);
        excelExport.export(new FileOutputStream("test.xlsx"));
    }

    @Test
    void 有夸多列的单元格单元格() throws IOException {
        ExcelExport excelExport = new ExcelExport();
        ExcelRow row = new ExcelRow();
        row.setHeight(30);
        excelExport.addRow(row);
        ExcelCell cell = new ExcelCell("这是一个单元格,很长很长很长", 200, 30);
        row.addCell(cell);
        excelExport.addMergeRange(new MergeRange(0, 0, 0, 2));
        excelExport.export(new FileOutputStream("test.xlsx"));
    }
    @AfterEach
    public void tearDown() {
        new File("test.xlsx").delete();
    }
}
