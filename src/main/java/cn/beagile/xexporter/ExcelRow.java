package cn.beagile.xexporter;


import java.util.ArrayList;
import java.util.List;

public class ExcelRow {
    private List<ExcelCell> cells = new ArrayList<>();
    private int height = 14;

    public List<ExcelCell> getCells() {
        return cells;
    }

    public int getHeight() {
        return height;
    }

    public void addCell(ExcelCell cell) {
        cells.add(cell);
    }

    public void setCells(List<ExcelCell> cells) {
        this.cells = cells;
    }

    public void setHeight(int height) {
        this.height = height;
    }

}
