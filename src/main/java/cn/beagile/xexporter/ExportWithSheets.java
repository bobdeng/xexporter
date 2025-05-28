package cn.beagile.xexporter;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class ExportWithSheets {
    private List<ExcelSheet> sheets;

    public List<ExcelSheet> getSheets() {
        return sheets;
    }

    public void setSheets(List<ExcelSheet> sheets) {
        this.sheets = sheets;
    }

    public void export(OutputStream outputStream) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        sheets.forEach(excelSheet -> {
            try {
                excelSheet.export(workbook);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });
        workbook.write(outputStream);
    }
}
