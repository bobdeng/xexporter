package cn.beagile.xexporter;

import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

@Data
public class ExcelExport {
    private List<ExcelRow> rows = new ArrayList<>();
    private List<MergeRange> mergeRanges = new ArrayList<>();

    public void addRow(ExcelRow row) {
        rows.add(row);
    }

    public void export(OutputStream outputStream) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        for (int i = 0; i < rows.size(); i++) {
            ExcelRow row = rows.get(i);
            org.apache.poi.xssf.usermodel.XSSFRow sheetRow = sheet.createRow(i);
            sheetRow.setHeightInPoints((short) row.getHeight());
            List<ExcelCell> cells = row.getCells();
            for (int j = 0; j < cells.size(); j++) {
                ExcelCell cell = cells.get(j);
                org.apache.poi.xssf.usermodel.XSSFCell sheetCell = sheetRow.createCell(j);
                CellStyle cellStyle = workbook.createCellStyle();
                XSSFFont font = workbook.createFont();
                font.setFontHeightInPoints((short) cell.getFontSize());
                if (cell.getFont() != null) {
                    font.setColor(org.apache.poi.ss.usermodel.IndexedColors.valueOf(cell.getFont().getColor()).getIndex());
                }
                cellStyle.setFont(font);
                BorderStyle border = BorderStyle.THIN;
                cellStyle.setBorderBottom(border);
                cellStyle.setBorderLeft(border);
                cellStyle.setBorderRight(border);
                cellStyle.setBorderTop(border);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setWrapText(true);
                sheetCell.setCellStyle(cellStyle);
                sheetCell.setCellValue(cell.getContent());

                int cellWidth = cell.getWidth() * 256;
                if (sheet.getColumnWidth(j) < cellWidth) {
                    sheet.setColumnWidth(j, cellWidth);
                }
            }
        }
        mergeRanges
                .stream().filter(MergeRange::needMerge)
                .forEach(mergeRange -> {
                    try {
                        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(mergeRange.getFirstRow(), mergeRange.getLastRow(), mergeRange.getFirstCol(), mergeRange.getLastCol()));
                    } catch (Exception e) {
                        System.out.println(mergeRange);
                    }
                });
        workbook.write(outputStream);
    }

    public void addMergeRange(MergeRange mergeRange) {
        mergeRanges.add(mergeRange);
    }
}
