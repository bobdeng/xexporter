package cn.beagile.xexporter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

public class ExportWithCells {
    private List<ExcelRow> rows = new ArrayList<>();
    private List<MergeRange> mergeRanges = new ArrayList<>();


    public void addRow(ExcelRow row) {
        rows.add(row);
    }

    public ExportWithCells() {
    }

    public void export(OutputStream outputStream) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();

        workbook.write(outputStream);
    }

    public void export(SXSSFWorkbook workbook, String sheetName) throws IOException {
        SXSSFSheet sheet = Optional.ofNullable(workbook.getSheet(sheetName))
                .orElseGet(() -> workbook.createSheet(sheetName));
        int start = sheet.getLastRowNum() + 1;
        for (int i = 0; i < rows.size(); i++) {
            ExcelRow row = rows.get(i);
            Row sheetRow = sheet.createRow(start + i);
            sheetRow.setHeightInPoints((short) row.getHeight());
            List<ExcelCell> cells = row.getCells();
            for (int j = 0; j < cells.size(); j++) {
                ExcelCell cell = cells.get(j);
                Cell sheetCell = sheetRow.createCell(j);
                if (cell.isNumber()) {
                    sheetCell.setCellValue(cell.doubleValue());
                    sheetCell.setCellType(org.apache.poi.ss.usermodel.CellType.NUMERIC);
                } else {
                    sheetCell.setCellValue(cell.getContent());
                    sheetCell.setCellType(org.apache.poi.ss.usermodel.CellType.STRING);
                }
                int cellWidth = cell.getWidth() * 256;
                if (sheet.getColumnWidth(j) < cellWidth) {
                    sheet.setColumnWidth(j, cellWidth);
                }
                if (cell.getFont() != null) {
                    XSSFFont font = (XSSFFont) workbook.createFont();
                    font.setFontHeightInPoints((short) cell.getFontSize());
                    font.setColor(IndexedColors.valueOf(cell.getFont().getColor()).index);
                    XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
                    style.setFont(font);
                    sheetCell.setCellStyle(style);
                }
            }
            if (i % 1000 == 0) {
                sheet.flushRows(1000);
            }
        }
        sheet.flushRows(1000);
        mergeRanges
                .stream().filter(MergeRange::needMerge)
                .forEach(mergeRange -> {
                    try {
                        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(mergeRange.getFirstRow(), mergeRange.getLastRow(), mergeRange.getFirstCol(), mergeRange.getLastCol()));
                    } catch (Exception e) {
                        System.out.println(mergeRange);
                    }
                });
    }

    public void addMergeRange(MergeRange mergeRange) {
        mergeRanges.add(mergeRange);
    }

    public List<ExcelRow> getRows() {
        return rows;
    }

    public void setRows(List<ExcelRow> rows) {
        this.rows = rows;
    }

    public List<MergeRange> getMergeRanges() {
        return mergeRanges;
    }

    public void setMergeRanges(List<MergeRange> mergeRanges) {
        this.mergeRanges = mergeRanges;
    }
}
