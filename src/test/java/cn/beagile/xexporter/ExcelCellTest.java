package cn.beagile.xexporter;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ExcelCellTest {
    @Test
    void parse_double(){
        var cell = new ExcelCell();
        cell.setContent("123,456.789");
        cell.setType("number");
        assertEquals(123456.789, cell.doubleValue(), 0.0001);
    }
    @Test
    void parse_percent(){
        var cell = new ExcelCell();
        cell.setContent("12.00%");
        cell.setType("percent");
        assertEquals(0.12, cell.doubleValue(), 0.0001);
    }
}
