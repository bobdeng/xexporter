package cn.beagile.xexporter;


public class MergeRange {
    private  int firstRow;
    private  int lastRow;
    private  int firstCol;
    private  int lastCol;

    public MergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {

        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;
    }

    public boolean needMerge() {
        return firstRow != lastRow || firstCol != lastCol;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public int getLastRow() {
        return lastRow;
    }

    public int getFirstCol() {
        return firstCol;
    }

    public int getLastCol() {
        return lastCol;
    }

    public void setFirstRow(int firstRow) {
        this.firstRow = firstRow;
    }

    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    public void setFirstCol(int firstCol) {
        this.firstCol = firstCol;
    }

    public void setLastCol(int lastCol) {
        this.lastCol = lastCol;
    }

    public MergeRange() {
    }
}
