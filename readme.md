# XExporter - Excel 导出工具库

XExporter 是一个功能强大的 Java Excel 导出工具库，基于 Apache POI，提供了灵活的 Excel 文件生成和模板填充功能。

## 特性

- ✅ **模板导出**：支持基于 Excel 模板的数据填充
- ✅ **动态数组**：自动展开数组数据，支持多层嵌套
- ✅ **单元格操作**：直接通过代码创建和配置单元格
- ✅ **多 Sheet 支持**：支持创建和管理多个工作表
- ✅ **图片插入**：支持在单元格中插入图片，自动调整大小
- ✅ **公式支持**：保留并自动更新 Excel 公式
- ✅ **样式保留**：保持模板中的单元格样式（字体、颜色、格式等）
- ✅ **合并单元格**：支持单元格合并操作
- ✅ **大数据量**：使用流式 API 支持大数据量导出

## 安装

### Maven
```xml
<repositories>
    <repository>
        <id>jitpack.io</id>
        <url>https://jitpack.io</url>
    </repository>
</repositories>

<dependency>
    <groupId>com.github.bobdeng</groupId>
    <artifactId>xexporter</artifactId>
    <version>1.14</version>
</dependency>
```

### Gradle
```groovy
repositories {
    maven { url 'https://jitpack.io' }
}

dependencies {
    implementation 'com.github.bobdeng:xexporter:1.14'
}
```

## 快速开始

### 1. 基于模板导出

使用 Excel 模板文件，通过占位符语法填充数据。

#### 模板语法

占位符使用 `≮` 和 `≯` 包裹，支持 JSONPath 表达式：

- **简单值**：`≮name≯` - 填充单个字段
- **嵌套对象**：`≮user.address.city≯` - 访问嵌套属性
- **数组元素**：`≮list[0].name≯` - 访问数组元素
- **数组展开**：`≮list[].name≯` - 自动展开数组，为每个元素创建新行
- **特殊字符**：`≮params['学院及专业班级']≯` - 访问包含特殊字符的键

#### 示例代码

```java
import cn.beagile.xexporter.ExportWithTemplate;
import com.google.gson.Gson;
import java.io.*;

// 准备数据
String json = """
{
    "name": "海洋学院",
    "city": "泉州",
    "age": 18,
    "list": [
        {"name": "张三", "age": 18},
        {"name": "李四", "age": 19}
    ]
}
""";

// 创建导出器
ExportWithTemplate exporter = new Gson().fromJson(json, ExportWithTemplate.class);

// 导出到文件
try (InputStream template = new FileInputStream("template.xlsx");
     OutputStream output = new FileOutputStream("output.xlsx")) {
    exporter.export(template, output);
}
```

#### 模板示例

**template.xlsx 内容：**
```
| 学院名称：≮name≯ | 城市：≮city≯ |
| 姓名          | 年龄        |
| ≮list[].name≯ | ≮list[].age≯ |
```

**导出结果：**
```
| 学院名称：海洋学院 | 城市：泉州 |
| 姓名    | 年龄 |
| 张三    | 18   |
| 李四    | 19   |
```

### 2. 直接创建单元格

不使用模板，直接通过代码创建 Excel 内容。

```java
import cn.beagile.xexporter.*;
import java.io.*;

ExportWithCells exporter = new ExportWithCells();

// 创建行
ExcelRow row = new ExcelRow();
row.setHeight(50);

// 创建单元格
ExcelCell cell = new ExcelCell("这是内容", 30, 14);
cell.setBgColor("GREEN");
cell.setFont(new ExcelCell.Font("RED"));
cell.setType("number"); // 设置为数字类型

row.addCell(cell);
exporter.addRow(row);

// 添加合并单元格
exporter.addMergeRange(new MergeRange(0, 0, 0, 2)); // 合并第1行的前3列

// 导出
try (OutputStream output = new FileOutputStream("output.xlsx")) {
    exporter.export(output);
}
```

### 3. 多 Sheet 导出

创建包含多个工作表的 Excel 文件。

```java
import cn.beagile.xexporter.*;
import java.io.*;
import java.util.List;

ExportWithSheets exporter = new ExportWithSheets();

// 创建第一个 Sheet
ExcelSheet sheet1 = new ExcelSheet();
sheet1.setName("销售数据");
sheet1.setActive(true); // 设置为活动 Sheet
sheet1.setCells(createCellsForSheet1());

// 创建第二个 Sheet
ExcelSheet sheet2 = new ExcelSheet();
sheet2.setName("统计报表");
sheet2.setCells(createCellsForSheet2());

exporter.setSheets(List.of(sheet1, sheet2));

// 导出
try (OutputStream output = new FileOutputStream("output.xlsx")) {
    exporter.export(output);
}
```

### 4. 插入图片

在模板中插入图片，支持自动调整大小和多图片并排显示。

```java
import cn.beagile.xexporter.*;
import com.google.gson.Gson;
import java.io.*;

String json = """
{
    "data": {
        "list": [
            {
                "name": "张三",
                "idCardPic": "images://photo1.jpg,photo2.jpg"
            }
        ]
    }
}
""";

ExportWithTemplate exporter = new Gson().fromJson(json, ExportWithTemplate.class);

// 设置文件读取器
exporter.setFileReader(path -> new File("/path/to/images/" + path));

try (InputStream template = new FileInputStream("template.xlsx");
     OutputStream output = new FileOutputStream("output.xlsx")) {
    exporter.export(template, output);
}
```

**图片语法：**
- `images://path1.jpg` - 插入单张图片
- `images://path1.jpg,path2.jpg,path3.jpg` - 插入多张图片（自动并排显示）
- 支持 JPG 和 PNG 格式
- 自动根据单元格大小调整图片尺寸
- 支持合并单元格中的图片插入

## 高级功能

### 数字和百分比格式

```java
ExcelCell cell = new ExcelCell();
cell.setContent("123,456.789");
cell.setType("number"); // 数字类型，自动去除千分位
// cell.doubleValue() 返回 123456.789

ExcelCell percentCell = new ExcelCell();
percentCell.setContent("12.00%");
percentCell.setType("percent"); // 百分比类型
// percentCell.doubleValue() 返回 0.12
```

### 保留单元格格式

模板中的单元格格式（货币、日期、百分比等）会自动保留：

```java
// 模板中如果单元格设置了货币格式（如 ¥1,234.56）
// 填充数字数据时会自动应用该格式
```

### 公式自动更新

模板中的公式在数组展开时会自动更新行号：

```java
// 模板中的公式：=SUM(A2:B2)
// 展开后自动更新为：=SUM(A3:B3), =SUM(A4:B4), ...
```

### 大数据量导出

使用流式 API 处理大量数据，避免内存溢出：

```java
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

SXSSFWorkbook workbook = new SXSSFWorkbook(100); // 内存中保留100行
ExportWithCells exporter = new ExportWithCells();

// 添加大量数据
for (int i = 0; i < 100000; i++) {
    ExcelRow row = new ExcelRow();
    row.addCell(new ExcelCell("数据" + i, 20, 10));
    exporter.addRow(row);
}

exporter.export(workbook, "数据表");
workbook.write(new FileOutputStream("large.xlsx"));
workbook.close();
```

## API 文档

### ExportWithTemplate

基于模板的导出器。

**方法：**
- `export(InputStream template, OutputStream output)` - 导出 Excel
- `setFileReader(FileReader reader)` - 设置文件读取器（用于图片）
- `readStringFromJson(String path)` - 从数据中读取值

### ExportWithCells

直接创建单元格的导出器。

**方法：**
- `addRow(ExcelRow row)` - 添加行
- `addMergeRange(MergeRange range)` - 添加合并单元格
- `export(OutputStream output)` - 导出到输出流
- `export(SXSSFWorkbook workbook, String sheetName)` - 导出到指定工作簿

### ExcelCell

单元格对象。

**属性：**
- `content` - 单元格内容
- `width` - 列宽（字符数）
- `fontSize` - 字体大小
- `bgColor` - 背景颜色
- `font` - 字体对象（包含颜色）
- `type` - 数据类型（"string", "number", "percent"）

**方法：**
- `doubleValue()` - 获取数字值
- `isNumber()` - 是否为数字类型

### ExcelRow

行对象。

**属性：**
- `height` - 行高（像素）
- `cells` - 单元格列表

### ExportWithSheets

多 Sheet 导出器。

**方法：**
- `setSheets(List<ExcelSheet> sheets)` - 设置工作表列表
- `export(OutputStream output)` - 导出到输出流

### ExcelSheet

工作表对象。

**属性：**
- `name` - Sheet 名称
- `active` - 是否为活动 Sheet
- `cells` - 单元格数据（ExportWithCells 对象）

## 测试

运行测试：

```bash
./gradlew test
```

查看测试报告：

```bash
open build/reports/tests/test/index.html
```

## 依赖

- Apache POI 5.4.0
- Gson 2.10
- JsonPath 2.9.0
- Java 17+

## 许可证

本项目使用 Apache License 2.0 许可证。

## 贡献

欢迎提交 Issue 和 Pull Request！

## 更新日志

### v1.14
- 支持图片插入功能
- 优化数字格式处理
- 修复百分比千分位显示问题
- 改进合并单元格支持

## 联系方式

- GitHub: https://github.com/bobdeng/xexporter
- JitPack: https://jitpack.io/#bobdeng/xexporter
