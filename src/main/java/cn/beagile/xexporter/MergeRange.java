package cn.beagile.xexporter;

import lombok.Data;
import lombok.ToString;

@Data
@ToString
public class MergeRange {
    private final int firstRow;
    private final int lastRow;
    private final int firstCol;
    private final int lastCol;

    public MergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {

        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;
    }

    public boolean needMerge() {
        return firstRow != lastRow || firstCol != lastCol;
    }
}
