package cn.beagile.xexporter;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
            sheets.forEach(excelSheet -> {
                try {
                    excelSheet.export(workbook);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            });
            for (int i = 0; i < sheets.size(); i++) {
                if (sheets.get(i).isActive()) {
                    workbook.setActiveSheet(i);
                }
            }
            workbook.write(outputStream);
        }
    }

    public void write(SXSSFWorkbook workbook) throws IOException {
        sheets.forEach(excelSheet -> {
            try {
                excelSheet.export(workbook);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });
        for (int i = 0; i < sheets.size(); i++) {
            if (sheets.get(i).isActive()) {
                workbook.setActiveSheet(i);
            }
        }
    }
}
