package cn.beagile.xexporter;


public class ExcelCell {
    private String content;
    private int width = 10;
    private int fontSize = 14;
    private Font font;
    private String bgColor;

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

    public static class Font {
        private String color;

        public Font() {
        }

        public Font(String color) {
            this.color = color;
        }

        public String getColor() {
            return color;
        }

        public void setColor(String color) {
            this.color = color;
        }
    }
}
