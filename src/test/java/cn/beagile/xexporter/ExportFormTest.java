package cn.beagile.xexporter;

import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertEquals;

class ExportFormTest {
    private ExportForm exportForm;
    private static String tempFile = "temp.xlsx";

    @BeforeEach
    public void setup() {
        String json = """
                {
                      "name": "导出数据到Excel",
                      "template":"测试导出",
                      "data":{
                        "recruitment":{
                            "params":{
                                "学院及专业班级":"海洋学院1"
                            }
                        },
                        "name":"海洋学院",
                        "city":"泉州",
                        "list":[{"name":"张三","age":18,"group":{"name":"班级1"}},{"name":"李四","age":19,"group":{"name":"班级2"}}]
                      }
                    }
                """;
        exportForm = new Gson().fromJson(json, ExportForm.class);
    }

    @AfterEach
    public void tearDown() {
        new File(tempFile).delete();
    }

    @Test
    public void get_array_data_value() {
        assertEquals("张三", exportForm.readStringFromJson("list[0].name"));
        assertEquals("李四", exportForm.readStringFromJson("list[1].name"));
        assertEquals("", exportForm.readStringFromJson("list[1].notExist"));
        assertEquals("海洋学院", exportForm.readStringFromJson("name"));
        assertEquals("海洋学院1", exportForm.readStringFromJson("recruitment.params['学院及专业班级']"));
        assertEquals("海洋学院1", exportForm.readStringFromJson("recruitment.params.学院及专业班级"));
    }

    @Test
    public void should_export_to_excel() throws IOException {
        exportForm.setTemplate("测试空");
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        exportForm.append(outputStream);
        outputStream.close();
        assertExcelContent("");
    }

    @Test
    public void should_export_one() throws IOException {
        exportForm.setTemplate("测试一个值");
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        exportForm.append(outputStream);
        outputStream.close();
        assertExcelContent("海洋学院");
    }

    @Test
    public void should_export_deep_array() throws IOException {
        exportForm.setTemplate("测试获取深度字段");
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        exportForm.append(outputStream);
        outputStream.close();
        assertExcelContent("班级1\n" +
                "班级2");
    }


    @Test
    public void should_export_array() throws IOException {
        exportForm.setTemplate("测试一个数组值");
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        exportForm.append(outputStream);
        outputStream.close();
        assertExcelContent("张三\n李四");
    }

    @Test
    public void should_export_2array() throws IOException {
        exportForm.setTemplate("测试2个数组值");
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        exportForm.append(outputStream);
        outputStream.close();
        assertExcelContent("""
                姓名,年龄,备注
                张三,18.0,测试
                李四,18.0,测试
                               
                年龄
                18.0
                19.0
                """);
    }

    @Test
    public void should_export_mix() throws IOException {
        exportForm.setTemplate("测试混合");
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        exportForm.append(outputStream);
        outputStream.close();
        assertExcelContent("""
                ,名称：海洋学院2泉州
                ,姓名,年龄
                ,张三,18.0
                ,李四,19.0
                """);
    }

    private void assertExcelContent(String expectContent) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(tempFile));
        XSSFSheet sheet = workbook.getSheetAt(0);
        String sheetContent = "";
        for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
            XSSFRow row = sheet.getRow(i);
            sheetContent += getRowString(row);
            sheetContent += "\n";
        }
        assertEquals(expectContent.trim(), sheetContent.trim());
    }

    private String getRowString(Row row) {
        if (row == null) {
            return "";
        }
        return IntStream.range(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .map(this::getCellString)
                .collect(Collectors.joining(","));

    }

    private String getCellString(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue() + "";
        }
        return "";
    }
    @Test
    void 重建公式() {
        String formula = "IF(LEN($A2)=18,MID($A2,7,4)&-MID($A2,11,2)&-MID($A2,13,2),(IF(LEN($A2)=15,19&MID($A2,7,2)&-MID($A2,9,2)&-MID($A2,11,2),\"\")))";
        String newFormula = exportForm.rebuildFormula(formula, 1, 19);
        assertThat(newFormula).isEqualTo("IF(LEN($A20)=18,MID($A20,7,4)&-MID($A20,11,2)&-MID($A20,13,2),(IF(LEN($A20)=15,19&MID($A20,7,2)&-MID($A20,9,2)&-MID($A20,11,2),\"\")))");
    }

}
