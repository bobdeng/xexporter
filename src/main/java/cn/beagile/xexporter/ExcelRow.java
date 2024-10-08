package cn.beagile.xexporter;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class ExcelRow {
    private List<ExcelCell> cells = new ArrayList<>();
    private int height = 14;

    public void addCell(ExcelCell cell) {
        cells.add(cell);
    }

}
