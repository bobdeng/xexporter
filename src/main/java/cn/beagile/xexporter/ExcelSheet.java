package cn.beagile.xexporter;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.List;

public class ExcelSheet {
    private String name;
    private ExportWithCells cells;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ExportWithCells getCells() {
        return cells;
    }

    public void setCells(ExportWithCells cells) {
        this.cells = cells;
    }

    public void export(XSSFWorkbook workbook) throws IOException {
        cells.export(workbook, name);
    }
}
