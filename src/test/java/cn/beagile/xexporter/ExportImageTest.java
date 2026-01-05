package cn.beagile.xexporter;

import com.google.common.io.Resources;
import com.google.gson.Gson;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

class ExportImageTest {
    private ExportWithTemplate exportForm;
    private static String tempFile = "temp.xlsx";

    @BeforeEach
    public void setup() {
        String json = """
                {
                  "template": "export_with_image",
                  "data": {
                    "list": [
                      {
                        "id": "200",
                        "publishImages":"images://",
                        "student": {
                          "id": "2",
                          "idCardPic":"images://123456,223456",
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
                         "idCardPic":"images://123456",
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
        exportForm.setFileReader(path -> new File("src/test/resources/test.jpg"));
    }

    @AfterEach
    public void tearDown() {
//        new File(tempFile).delete();
    }

    @Test
    public void export() throws IOException {
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        ByteArrayInputStream templateInputStream = new ByteArrayInputStream(Resources.toByteArray(Resources.getResource("template/export_with_image.xlsx")));
        exportForm.export(templateInputStream, outputStream);
    }

}
