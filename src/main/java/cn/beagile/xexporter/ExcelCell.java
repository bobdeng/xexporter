package cn.beagile.xexporter;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelCell {
    private String content;
    private int width = 10;
    private int fontSize = 14;
    private Font font;

    public ExcelCell(String content, int width, int fontSize) {
        this.content = content;
        this.width = width;
        this.fontSize = fontSize;
    }

    @Data
    public static class Font {
        private String color;
    }
}
