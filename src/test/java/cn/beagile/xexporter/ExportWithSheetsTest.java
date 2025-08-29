package cn.beagile.xexporter;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

class ExportWithSheetsTest {
    @Test
    public void export_with_sheets() throws IOException {
        ExportWithSheets sheets = new ExportWithSheets();
        sheets.setSheets(List.of(getExcelSheet("123", false), getExcelSheet("234", true)));
        sheets.export(new FileOutputStream("test.xlsx"));
    }

    @Test
    public void export_with_sheets_append() throws IOException {
        ExportWithSheets sheets = new ExportWithSheets();
        sheets.setSheets(List.of(getExcelSheet("123", false), getExcelSheet("234", true)));
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        sheets.write(workbook);
        sheets.write(workbook);
        workbook.write(new FileOutputStream("test.xlsx"));
        workbook.close();
    }

    private static ExcelSheet getExcelSheet(String name, boolean active) {
        ExportWithCells excelExport = new ExportWithCells();
        ExcelRow row = new ExcelRow();
        row.setHeight(50);
        excelExport.addRow(row);
        ExcelCell cell = new ExcelCell("这是一个单元格,,很长很长很长", 30, 14);
        cell.setFont(new ExcelCell.Font("RED"));
        cell.setBgColor("GREEN");
        row.addCell(cell);
        ExcelSheet sheet1 = new ExcelSheet();
        sheet1.setName(name);
        sheet1.setActive(active);
        sheet1.setCells(excelExport);
        return sheet1;
    }
}
