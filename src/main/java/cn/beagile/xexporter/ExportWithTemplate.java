package cn.beagile.xexporter;

import com.jayway.jsonpath.JsonPath;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Objects;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

public class ExportWithTemplate {
    private Object data;
    private String excelType;
    private FileReader fileReader;

    public void setFileReader(FileReader fileReader) {
        this.fileReader = fileReader;
    }

    public ExportWithTemplate(Object data, String excelType) {
        this.data = data;
        this.excelType = excelType;
    }

    public ExportWithTemplate(Object data, String excelType, FileReader fileReader) {
        this.data = data;
        this.excelType = excelType;
        this.fileReader = fileReader;
    }

    public ExportWithTemplate() {
    }


    private static final Pattern placeholderPattern = Pattern.compile("≮[^≯]*≯");

    public void export(InputStream templateInputStream, OutputStream outputStream) throws IOException {
        Workbook workbook = readWorkbookFromTemplate(templateInputStream);
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

    private Workbook readWorkbookFromTemplate(InputStream templateInputStream) {
        Supplier<Workbook> xlsxWorkbook = () -> getXlsxWorkbook(templateInputStream);
        Supplier<Workbook> xlsWorkbook = () -> getXlsWorkbook(templateInputStream);
        return isXlsx() ? xlsxWorkbook.get() : xlsWorkbook.get();
    }

    private boolean isXlsx() {
        if (excelType == null) {
            return true;
        }
        return "xlsx".equals(excelType);
    }

    private Workbook getXlsxWorkbook(InputStream templateInputStream) {
        try {
            return new XSSFWorkbook(templateInputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private Workbook getXlsWorkbook(InputStream templateInputStream) {
        try {
            return new HSSFWorkbook(templateInputStream);
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
        return JsonPath.read(this.data, jsonPath);
    }

    public String getSingleValueFromJson(String name) {
        return readStringFromJson(removeDecoration(name));
    }

    private String removeDecoration(String name) {
        return name.substring(1, name.length() - 1);
    }

    public String readStringFromJson(String name) {
        try {
            return JsonPath.read(this.data, "$." + name).toString();
        } catch (Exception e) {
            return "";
        }
    }

    private void fillRow(Row row) {
        IntStream.rangeClosed(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .filter(Objects::nonNull)
                .filter(this::isPlaceholder)
                .forEach(this::fillCell);
    }

    private void fillCell(Cell cell) {
        // 保存原始的单元格样式
        CellStyle originalStyle = cell.getCellStyle();

        String cellValue = cell.getStringCellValue();
        while (isPlaceholder(cellValue)) {
            cellValue = run(cellValue);
        }

        // 检查是否是图片路径
        if (cellValue.startsWith("images://")) {
            handleImages(cell, cellValue);
            return;
        }

        // 尝试将值转换为合适的类型并设置
        setCellValueWithTypeDetection(cell, cellValue);

        // 恢复原始样式
        cell.setCellStyle(originalStyle);
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

        // 保持原来的行高
        newRow.setHeight(row.getHeight());

        copyRowCells(rowIndex, offset, row, newRow);

        copyArrayCells(row, newRow, offset);

        // 复制合并单元格
        copyMergedRegions(sheet, rowIndex, rowIndex + offset + 1);
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
        Cell targetCell = target.getCell(cellIndex);
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

    private void setCellValueWithTypeDetection(Cell cell, String value) {
        if (value == null || value.trim().isEmpty()) {
            cell.setCellValue("");
            return;
        }
        cell.setCellValue(value);
    }

    /**
     * 处理 images:// 前缀的图片路径
     *
     * @param cell      目标单元格
     * @param cellValue 格式: images://path1,path2,path3
     */
    private void handleImages(Cell cell, String cellValue) {
        if (fileReader == null) {
            cell.setCellValue("FileReader not configured");
            return;
        }

        // 移除 images:// 前缀
        String pathsString = cellValue.substring("images://".length());

        // 按逗号分割路径
        String[] paths = pathsString.split(",");

        // 过滤空路径
        int imageCount = 0;
        for (String path : paths) {
            if (path.trim().isEmpty()) {
                continue;
            }
            imageCount++;
        }

        if (imageCount == 0) {
            return;
        }

        // 清空单元格内容
        cell.setCellValue("");

        // 获取工作表和工作簿
        Sheet sheet = cell.getSheet();
        Workbook workbook = sheet.getWorkbook();
        Drawing<?> drawing = sheet.createDrawingPatriarch();

        // 1. 读取单元格宽高
        int rowIndex = cell.getRowIndex();
        int colIndex = cell.getColumnIndex();
        Row row = cell.getRow();

        // 检查是否是合并单元格
        org.apache.poi.ss.util.CellRangeAddress mergedRegion = getMergedRegion(cell);

        int cellWidthPixels;
        int cellHeightPixels;

        if (mergedRegion != null) {
            // 如果是合并单元格，计算合并区域的总宽高
            cellWidthPixels = getMergedRegionWidth(sheet, mergedRegion);
            cellHeightPixels = getMergedRegionHeight(sheet, mergedRegion);
        } else {
            // 如果不是合并单元格，计算单个单元格的宽高
            // Excel 列宽单位是字符宽度，行高单位是点(point)
            int cellWidth = sheet.getColumnWidth(colIndex); // 单位是 1/256 字符宽度
            float cellHeight = row.getHeightInPoints(); // 单位是点

            // 转换为像素（大约）
            // 列宽：1 字符 ≈ 7 像素
            cellWidthPixels = (int) (cellWidth / 256.0 * 7);
            // 行高：1 点 ≈ 1.33 像素
            cellHeightPixels = (int) (cellHeight * 1.33);
        }

        // 2. 宽高减去10%
        double marginRatio = 0.9; // 保留90%，减去10%
        int usableWidth = (int) (cellWidthPixels * marginRatio);
        int usableHeight = (int) (cellHeightPixels * marginRatio);

        // 3. 根据图片数量计算图片最大宽高
        int maxImageWidth = usableWidth / imageCount;
        int maxImageHeight = usableHeight;

        int currentImageIndex = 0;
        for (String path : paths) {
            path = path.trim();
            if (path.isEmpty()) {
                continue;
            }

            try {
                // 使用 fileReader 读取图片文件
                File imageFile = fileReader.read(path);
                if (imageFile == null || !imageFile.exists()) {
                    continue;
                }

                // 读取图片以获取原始尺寸
                BufferedImage bufferedImage = ImageIO.read(imageFile);
                if (bufferedImage == null) {
                    throw new RuntimeException("图片格式不支持");
                }

                int originalWidth = bufferedImage.getWidth();
                int originalHeight = bufferedImage.getHeight();

                // 4. 根据图片大小计算宽高，不要超过步骤3计算的宽高
                // 计算宽度和高度的缩放比例
                double widthScale = (double) maxImageWidth / originalWidth;
                double heightScale = (double) maxImageHeight / originalHeight;

                // 取最小的缩放比例，确保图片不超出最大宽高限制
                double scale = Math.min(widthScale, heightScale);

                // 按照最小缩放比例计算缩放后的尺寸
                int scaledWidth = (int) (originalWidth * scale);
                int scaledHeight = (int) (originalHeight * scale);

                // 读取图片字节
                byte[] imageBytes;
                try (FileInputStream fis = new FileInputStream(imageFile)) {
                    imageBytes = IOUtils.toByteArray(fis);
                }

                // 确定图片类型
                int pictureType = getPictureType(imageFile.getName());
                if (pictureType == -1) {
                    continue;
                }

                // 添加图片到工作簿
                int pictureIdx = workbook.addPicture(imageBytes, pictureType);

                // 创建锚点
                ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();

                // 计算当前图片的位置
                // dx1, dy1, dx2, dy2 的单位是 EMU (English Metric Units)
                // 对于XSSF: 1像素 ≈ 9525 EMU
                int emuPerPixel = 9525;

                // 计算边距（总边距10%平均分配）
                int horizontalMargin = (cellWidthPixels - usableWidth) / 2;
                int verticalMargin = (cellHeightPixels - usableHeight) / 2;

                // 计算起始位置（每张图片占据平均宽度，并居中显示）
                int startXPixels = horizontalMargin + currentImageIndex * maxImageWidth + (maxImageWidth - scaledWidth) / 2;
                // 图片在可用高度内垂直居中
                int startYPixels = verticalMargin + (maxImageHeight - scaledHeight) / 2;

                // 计算图片右下角的位置（像素）
                int endXPixels = startXPixels + scaledWidth;
                int endYPixels = startYPixels + scaledHeight;

                // 根据像素位置计算实际的单元格和偏移量
                // 确定起始位置所在的单元格和偏移
                int startCol = colIndex;
                int startRow = rowIndex;
                int dx1 = startXPixels;
                int dy1 = startYPixels;

                if (mergedRegion != null) {
                    // 对于合并单元格，需要找到 startXPixels 落在哪一列
                    int accumulatedWidth = 0;
                    for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
                        int colWidthPixels = (int) (sheet.getColumnWidth(col) / 256.0 * 7);
                        if (accumulatedWidth + colWidthPixels > startXPixels) {
                            startCol = col;
                            dx1 = startXPixels - accumulatedWidth;
                            break;
                        }
                        accumulatedWidth += colWidthPixels;
                    }

                    // 找到 startYPixels 落在哪一行
                    int accumulatedHeight = 0;
                    for (int r = mergedRegion.getFirstRow(); r <= mergedRegion.getLastRow(); r++) {
                        Row tempRow = sheet.getRow(r);
                        int rowHeightPixels = 0;
                        if (tempRow != null) {
                            rowHeightPixels = (int) (tempRow.getHeightInPoints() * 1.33);
                        }
                        if (accumulatedHeight + rowHeightPixels > startYPixels) {
                            startRow = r;
                            dy1 = startYPixels - accumulatedHeight;
                            break;
                        }
                        accumulatedHeight += rowHeightPixels;
                    }
                }

                // 确定结束位置所在的单元格和偏移
                int endCol = colIndex;
                int endRow = rowIndex;
                int dx2 = endXPixels;
                int dy2 = endYPixels;

                if (mergedRegion != null) {
                    // 找到 endXPixels 落在哪一列
                    int accumulatedWidth = 0;
                    for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
                        int colWidthPixels = (int) (sheet.getColumnWidth(col) / 256.0 * 7);
                        if (accumulatedWidth + colWidthPixels > endXPixels) {
                            endCol = col;
                            dx2 = endXPixels - accumulatedWidth;
                            break;
                        }
                        accumulatedWidth += colWidthPixels;
                        // 如果到了最后一列还没找到，说明超出范围，使用最后一列
                        if (col == mergedRegion.getLastColumn()) {
                            endCol = col;
                            dx2 = colWidthPixels;
                        }
                    }

                    // 找到 endYPixels 落在哪一行
                    int accumulatedHeight = 0;
                    for (int r = mergedRegion.getFirstRow(); r <= mergedRegion.getLastRow(); r++) {
                        Row tempRow = sheet.getRow(r);
                        int rowHeightPixels = 0;
                        if (tempRow != null) {
                            rowHeightPixels = (int) (tempRow.getHeightInPoints() * 1.33);
                        }
                        if (accumulatedHeight + rowHeightPixels > endYPixels) {
                            endRow = r;
                            dy2 = endYPixels - accumulatedHeight;
                            break;
                        }
                        accumulatedHeight += rowHeightPixels;
                        // 如果到了最后一行还没找到，说明超出范围，使用最后一行
                        if (r == mergedRegion.getLastRow()) {
                            endRow = r;
                            dy2 = rowHeightPixels;
                        }
                    }
                }

                // 设置锚点
                anchor.setCol1(startCol);
                anchor.setRow1(startRow);
                anchor.setCol2(endCol);
                anchor.setRow2(endRow);

                anchor.setDx1(dx1 * emuPerPixel);
                anchor.setDy1(dy1 * emuPerPixel);
                anchor.setDx2(dx2 * emuPerPixel);
                anchor.setDy2(dy2 * emuPerPixel);

                // 插入图片
                Picture picture = drawing.createPicture(anchor, pictureIdx);

                currentImageIndex++;

            } catch (IOException e) {
                // 忽略单个图片的错误，继续处理下一张
                e.printStackTrace();
            }
        }
    }

    /**
     * 根据文件扩展名确定图片类型
     */
    private int getPictureType(String fileName) {
        String extension = fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
        return switch (extension) {
            case "jpg", "jpeg" -> Workbook.PICTURE_TYPE_JPEG;
            case "png" -> Workbook.PICTURE_TYPE_PNG;
            default -> -1;
        };
    }

    /**
     * 复制源行的合并单元格区域到目标行
     *
     * @param sheet     工作表
     * @param sourceRow 源行索引
     * @param targetRow 目标行索引
     */
    private void copyMergedRegions(Sheet sheet, int sourceRow, int targetRow) {
        // 获取所有合并单元格区域
        int numMergedRegions = sheet.getNumMergedRegions();

        // 遍历所有合并单元格区域
        for (int i = 0; i < numMergedRegions; i++) {
            org.apache.poi.ss.util.CellRangeAddress mergedRegion = sheet.getMergedRegion(i);

            // 检查合并区域是否涉及源行
            if (mergedRegion.getFirstRow() == sourceRow && mergedRegion.getLastRow() == sourceRow) {
                // 创建新的合并区域（同样的列范围，但在目标行）
                org.apache.poi.ss.util.CellRangeAddress newMergedRegion = new org.apache.poi.ss.util.CellRangeAddress(
                        targetRow,
                        targetRow,
                        mergedRegion.getFirstColumn(),
                        mergedRegion.getLastColumn()
                );

                // 添加新的合并区域
                sheet.addMergedRegion(newMergedRegion);
            }
        }
    }

    /**
     * 获取单元格所在的合并区域
     *
     * @param cell 单元格
     * @return 合并区域，如果不在合并区域中则返回 null
     */
    private org.apache.poi.ss.util.CellRangeAddress getMergedRegion(Cell cell) {
        Sheet sheet = cell.getSheet();
        int rowIndex = cell.getRowIndex();
        int colIndex = cell.getColumnIndex();

        int numMergedRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numMergedRegions; i++) {
            org.apache.poi.ss.util.CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.isInRange(rowIndex, colIndex)) {
                return mergedRegion;
            }
        }
        return null;
    }

    /**
     * 计算合并单元格区域的总宽度（像素）
     *
     * @param sheet         工作表
     * @param mergedRegion  合并区域
     * @return 总宽度（像素）
     */
    private int getMergedRegionWidth(Sheet sheet, org.apache.poi.ss.util.CellRangeAddress mergedRegion) {
        int totalWidth = 0;
        for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
            int colWidth = sheet.getColumnWidth(col); // 单位是 1/256 字符宽度
            totalWidth += (int) (colWidth / 256.0 * 7); // 转换为像素
        }
        return totalWidth;
    }

    /**
     * 计算合并单元格区域的总高度（像素）
     *
     * @param sheet         工作表
     * @param mergedRegion  合并区域
     * @return 总高度（像素）
     */
    private int getMergedRegionHeight(Sheet sheet, org.apache.poi.ss.util.CellRangeAddress mergedRegion) {
        float totalHeight = 0;
        for (int row = mergedRegion.getFirstRow(); row <= mergedRegion.getLastRow(); row++) {
            Row rowObj = sheet.getRow(row);
            if (rowObj != null) {
                totalHeight += rowObj.getHeightInPoints(); // 单位是点
            }
        }
        return (int) (totalHeight * 1.33); // 转换为像素
    }
}
