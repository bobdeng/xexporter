package cn.beagile.xexporter;

import com.google.common.io.Resources;
import com.google.gson.Gson;
import com.jayway.jsonpath.Configuration;
import com.jayway.jsonpath.JsonPath;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

@Data
public class ExportWithTemplate {
    private String name;
    private String template;
    private Map data;
    private Object document;
    private FillConfig config;
    private String type;
    private String excelType;
    private static final Pattern placeholderPattern = Pattern.compile("≮[^≯]*≯");

    public void write(OutputStream outputStream) throws IOException {
        if ("fill".equals(type)) {
            this.fill(outputStream);
            return;
        }
        this.append(outputStream);
    }

    public void append(OutputStream outputStream) throws IOException {
        Workbook workbook = readWorkbookFromTemplate();
        Sheet sheet = workbook.getSheetAt(0);
        expandAllArrayPlaceholders(sheet);
        fillAllPlaceholders(sheet);
        rebuildFormula(sheet);
        writeWorkbook(outputStream, workbook);
    }

    private void rebuildFormula(Sheet sheet) {
        FormulaEvaluator formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        IntStream.rangeClosed(0, sheet.getLastRowNum())
                .mapToObj(sheet::getRow)
                .filter(Objects::nonNull)
                .forEach(row -> IntStream.rangeClosed(0, row.getLastCellNum())
                        .mapToObj(row::getCell)
                        .filter(Objects::nonNull)
                        .filter(cell -> cell.getCellType().equals(CellType.FORMULA))
                        .forEach(cell -> {
                            try {
                                formulaEvaluator.evaluateFormulaCell(cell);
                            } catch (Exception e) {
                            }
                        }));
    }

    private static void writeWorkbook(OutputStream outputStream, Workbook workbook) throws IOException {
        workbook.write(outputStream);
    }

    private void fillAllPlaceholders(Sheet sheet) {
        IntStream.rangeClosed(0, sheet.getLastRowNum())
                .mapToObj(sheet::getRow)
                .filter(Objects::nonNull)
                .forEach(this::fillRow);
    }

    private void expandAllArrayPlaceholders(Sheet sheet) {
        while (isNeedExpand(sheet)) {
            expandArray(sheet);
        }
    }

    private Workbook readWorkbookFromTemplate() {
        Supplier<Workbook> xlsxWorkbook = this::getXlsxWorkbook;
        Supplier<Workbook> xlsWorkbook = this::getXlsWorkbook;
        return isXlsx() ? xlsxWorkbook.get() : xlsWorkbook.get();
    }

    private boolean isXlsx() {
        if (excelType == null) {
            return true;
        }
        return "xlsx".equals(excelType);
    }

    private Workbook getXlsxWorkbook() {
        try {
            return new XSSFWorkbook(new ByteArrayInputStream(Resources.toByteArray(Resources.getResource("template/" + template + ".xlsx"))));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private Workbook getXlsWorkbook() {
        try {
            byte[] byteArray = Resources.toByteArray(Resources.getResource("template/" + template + ".xls"));
            return new HSSFWorkbook(new ByteArrayInputStream(byteArray));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void expandArray(Sheet sheet) {
        IntStream.rangeClosed(0, sheet.getLastRowNum())
                .mapToObj(sheet::getRow)
                .filter(this::isArray)
                .forEach(row -> appendArray(sheet, row.getRowNum()));
    }

    private boolean isNeedExpand(Sheet sheet) {
        return IntStream.rangeClosed(0, sheet.getLastRowNum())
                .mapToObj(sheet::getRow)
                .filter(Objects::nonNull)
                .anyMatch(this::isArray);
    }

    private int getArrayLengthOfRow(Row row) {
        return IntStream.rangeClosed(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .filter(this::isArrayPlaceholder)
                .map(this::getArrayLength)
                .findFirst().orElse(0);
    }

    private int getArrayLength(Cell cell) {
        String name = removeDecoration(cell.getStringCellValue());
        name = name.substring(0, name.indexOf("[]"));
        String jsonPath = "$." + name + ".length()";
        return JsonPath.read(getDocument(), jsonPath);
    }

    public String getSingleValueFromJson(String name) {
        return readStringFromJson(removeDecoration(name));
    }

    private String removeDecoration(String name) {
        return name.substring(1, name.length() - 1);
    }

    public String readStringFromJson(String name) {
        try {
            return JsonPath.read(getDocument(), "$." + name).toString();
        } catch (Exception e) {
            return "";
        }
    }

    private Object getDocument() {
        if (document == null) {
            document = Configuration.defaultConfiguration().jsonProvider().parse(new Gson().toJson(data));
        }
        return document;
    }

    private void fillRow(Row row) {
        IntStream.rangeClosed(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .filter(Objects::nonNull)
                .filter(this::isPlaceholder)
                .forEach(this::fillCell);
    }

    private void fillCell(Cell cell) {
        String cellValue = cell.getStringCellValue();
        while (isPlaceholder(cellValue)) {
            cellValue = run(cellValue);
        }
        cell.setCellValue(cellValue);
    }

    public String run(String cellValue) {
        Matcher matcher = placeholderPattern.matcher(cellValue);
        if (matcher.find()) {
            String placeholder = matcher.group();
            String value = getSingleValueFromJson(placeholder);
            return cellValue.replace(placeholder, value);
        }
        return cellValue;
    }

    private boolean isPlaceholder(Cell cell) {
        if (!cell.getCellType().equals(CellType.STRING)) {
            return false;
        }
        String cellValue = cell.getStringCellValue();
        return isPlaceholder(cellValue);
    }

    private static boolean isPlaceholder(String cellValue) {
        Matcher matcher = placeholderPattern.matcher(cellValue);
        return matcher.find();
    }

    private boolean isArrayPlaceholder(Cell cell) {
        if (cell == null) {
            return false;
        }
        return isPlaceholder(cell) && cell.getStringCellValue().contains("[]");
    }

    private void appendArray(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        IntStream.range(0, getArrayLengthOfRow(row))
                .forEach(i -> appendRow(sheet, rowIndex, i));
        removeFirstRowAndShiftBelowIt(sheet, rowIndex, row);
    }

    private void appendRow(Sheet sheet, int rowIndex, int offset) {
        Row row = sheet.getRow(rowIndex);
        if (hasRowsBelow(sheet, rowIndex, offset)) {
            sheet.shiftRows(rowIndex + offset + 1, sheet.getLastRowNum(), 1, true, false);
        }
        //Row newRow = sheet.createRow(rowIndex + offset + 1);
        Row newRow = sheet.createRow(rowIndex + offset + 1);
        copyRowCells(rowIndex, offset, row, newRow);

        copyArrayCells(row, newRow, offset);
    }

    private void copyRowCells(int rowIndex, int offset, Row row, Row newRow) {
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            Cell newCell = newRow.createCell(i, cell.getCellType());
            if (newCell.getCellType().equals(CellType.FORMULA)) {
                newCell.setCellFormula(rebuildFormula(cell.getCellFormula(), rowIndex, rowIndex + offset + 1));
            }
        }
    }

    private boolean hasRowsBelow(Sheet sheet, int rowIndex, int offset) {
        return rowIndex + offset + 1 < sheet.getLastRowNum();
    }

    private static void removeFirstRowAndShiftBelowIt(Sheet sheet, int start, Row row) {
        sheet.removeRow(row);
        sheet.shiftRows(start + 1, sheet.getLastRowNum(), -1);
    }

    private boolean isArray(Row row) {
        if (row == null) {
            return false;
        }
        return IntStream.rangeClosed(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .anyMatch(this::isArrayPlaceholder);
    }

    private void copyArrayCells(Row source, Row target, int rowIndex) {
        for (int cellIndex = 0; cellIndex < source.getLastCellNum(); cellIndex++) {
            Cell cell = source.getCell(cellIndex);
            copyCell(target, rowIndex, cellIndex, cell);
        }
    }

    private void copyCell(Row target, int rowIndex, int cellIndex, Cell cell) {
        Cell targetCell = target.getCell(cellIndex);
        if (cell != null) {
            CellStyle style = cell.getCellStyle();
            targetCell.setCellStyle(style);
        }
        if (isArrayPlaceholder(cell)) {
            copyArrayCellAddIndex(target, rowIndex, cellIndex, cell);
            return;
        }
        justCopyCellContent(target, cellIndex, cell);
    }

    private void justCopyCellContent(Row target, int cellIndex, Cell cell) {
        if (cell == null) {
            return;
        }
        if (cell.getCellType().equals(CellType.FORMULA)) {
            return;
        }
        Cell targetCell = target.createCell(cellIndex);
        switch (cell.getCellType()) {
            case NUMERIC -> targetCell.setCellValue(cell.getNumericCellValue());
            case BOOLEAN -> targetCell.setCellValue(cell.getBooleanCellValue());
            case BLANK -> targetCell.setBlank();
            default -> targetCell.setCellValue(cell.getStringCellValue());
        }
    }

    private void copyArrayCellAddIndex(Row target, int rowIndex, int cellIndex, Cell cell) {
        Cell targetCell = target.getCell(cellIndex);
        targetCell.setCellValue(cell.getStringCellValue().replace("[]", "[" + rowIndex + "]"));
    }


    public void fill(OutputStream outputStream) throws IOException {
        Workbook workbook = readWorkbookFromTemplate();
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 0; i < 400; i++) {
            String name = getSingleValueFromJson("≮" + this.config.listName + "[" + i + "]." + this.config.columns.get(0).name + "≯");
            if (name.isEmpty()) {
                break;
            }
            for (FillColumn column : this.config.columns) {
                Row row = sheet.getRow(i + 1);
                Cell cell = row.getCell(column.index);
                String value = getSingleValueFromJson("≮" + this.config.listName + "[" + i + "]." + column.name + "≯");
                cell.setCellValue(value);
            }
        }
        writeWorkbook(outputStream, workbook);
    }

    public String rebuildFormula(String formula, int originRowIndex, int rowIndex) {
        String previousRowName = "" + (originRowIndex + 1);
        String rowName = "" + (rowIndex + 1);
        String regex = "([A-Z]+)" + previousRowName;
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(formula);
        while (matcher.find()) {
            String column = matcher.group(1);
            formula = matcher.replaceAll(column + rowName);
        }
        return formula;
    }

    @Data
    public static class FillConfig {
        private List<FillColumn> columns;
        private String listName;

        public FillConfig(List<FillColumn> columns, String listName) {
            this.columns = columns;
            this.listName = listName;
        }
    }

    @Data
    public static class FillColumn {
        private int index;
        private String name;

        public FillColumn(int index, String name) {
            this.index = index;
            this.name = name;
        }
    }
}
