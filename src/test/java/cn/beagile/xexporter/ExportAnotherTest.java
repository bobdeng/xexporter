package cn.beagile.xexporter;

import com.google.gson.Gson;
import lombok.SneakyThrows;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

class ExportAnotherTest {
    private ExportWithTemplate exportForm;
    private static String tempFile= "temp.xlsx";

    @BeforeEach
    public void setup() {
        String json = """
                {
                  "name": "202301期船长班级学员-三管副.xlsx",
                  "template": "三管副",
                  "data": {
                    "list": [
                      {
                        "id": "200",
                        "student": {
                          "id": "2",
                          "register": {
                            "name": "张三",
                            "idNumber": "122222",
                            "contactPhone": "12345678901",
                            "bundle": "A套餐",
                            "project": {
                              "id": 100,
                              "name": "船长"
                            },
                            "namePinyin": "ZhangSan"
                          },
                          "recruitment": {
                            "id": 100,
                            "name": "船员培训-2023",
                            "coursePlanStartAt": "2023-03-01",
                            "coursePlanEndAt": "2023-09-30"
                          }
                        },
                        "index": "1"
                      },
                      {
                        "id": "200",
                        "student": {
                          "id": "200",
                          "register": {
                            "name": "李四",
                            "idNumber": "122222",
                            "namePinyin": "LiSi"
                          },
                          "recruitment": {
                            "id": 100,
                            "name": "船员培训-2023"
                          }
                        },
                        "index": "2"
                      }
                    ]
                  }
                }
                """;
        exportForm = new Gson().fromJson(json, ExportWithTemplate.class);
    }

    @AfterEach
    public void tearDown() {
        new File(tempFile).delete();
    }

    @SneakyThrows
    @Test
    public void export() throws FileNotFoundException {
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        exportForm.append(outputStream);
    }

}
