package cn.beagile.xexporter;


import java.text.NumberFormat;

public class ExcelCell {
    private String content;
    private int width = 10;
    private int fontSize = 9;
    private Font font;
    private String bgColor;
    private String type = "string"; // default type is string

    public boolean isNumber() {
        return "number".equals(type) || "percent".equals(type);
    }

    public ExcelCell() {
    }

    public ExcelCell(String content, int width, int fontSize) {
        this.content = content;
        this.width = width;
        this.fontSize = fontSize;
    }

    public String getContent() {
        return content;
    }

    public int getWidth() {
        return width;
    }

    public int getFontSize() {
        return fontSize;
    }

    public Font getFont() {
        return font;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public String getBgColor() {
        return bgColor;
    }

    public void setBgColor(String bgColor) {
        this.bgColor = bgColor;
    }

    public void setFont(Font font) {
        this.font = font;
    }

    public double doubleValue() {
        if (content == null || content.isEmpty()) {
            return 0;
        }
        try {
            if ("number".equals(type)) {
                return Double.parseDouble(content.replace(",", ""));
            }
            if ("percent".equals(type)) {
                return Double.parseDouble(content.replace("%", "")) / 100;
            }
            return Double.parseDouble(content.replace(",", ""));
        } catch (Exception e) {
            // If parsing fails, return 0
            return 0;
        }
    }

    public static class Font {
        private String color;

        public Font() {
        }

        public String getColor() {
            return color;
        }

        public Font(String color) {
            this.color = color;
        }
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }
}
