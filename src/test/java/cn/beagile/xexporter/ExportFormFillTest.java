package cn.beagile.xexporter;

import com.google.common.io.Resources;
import com.google.gson.Gson;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

class ExportFormFillTest {
    @Test
    void 导出Excel97() throws IOException {
        String json = """
                {
                  "template": "三管副学历生",
                  "excelType": "xls",
                  "data": {
                    "list": [
                      {
                        "id": "200",
                        "idNumber":"640302196301122294",
                        "student": {
                          "id": "2",
                          "register": {
                            "name": "张三",
                            "idNumber": "640302196301122294",
                            "contactPhone": "12345678901",
                            "bundle": "A套餐",
                            "project": {
                              "id": 100,
                              "name": "船长"
                            },
                            "namePinyin": "ZhangSan"
                          },
                          "exam": {
                              "examType": "申考适任",
                              "dutyType": "船长"
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
                        "idNumber": "140221198509097336",
                        "student": {
                        "exam": {
                              "examType": "申考适任",
                              "dutyType": "船长"
                            },
                          "id": "200",
                          "register": {
                            "name": "李四",
                            "idNumber": "140221198509097336",
                            "namePinyin": "LiSi"
                          },
                          "recruitment": {
                            "id": 100,
                            "name": "船员培训-2023"
                          }
                        },
                        "index": "2"
                      },
                      {
                        "id": "200",
                         "idNumber":"640302196301122294",
                        "student": {
                        "exam": {
                              "examType": "申考适任",
                              "dutyType": "船长"
                            },
                          "id": "200",
                          "register": {
                            "name": "李四",
                            "idNumber": "140221198509097336",
                            "namePinyin": "LiSi"
                          },
                          "recruitment": {
                            "id": 100,
                            "name": "船员培训-2023"
                          }
                        },
                        "index": "2"
                      },
                      {
                        "id": "200",
                         "idNumber":"640302196301122294",
                        "student": {
                        "exam": {
                              "examType": "申考适任",
                              "dutyType": "船长"
                            },
                          "id": "200",
                          "register": {
                            "name": "李四",
                            "idNumber": "140221198509097336",
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
                  },
                  "config": {
                    "listName": "list",
                    "columns": [
                      {
                        "index": 1,
                        "name": "student.register.name"
                      },{
                        "index": 0,
                        "name": "student.register.idNumber"
                      }
                
                    ]
                  }
                }
                """;
        ExportWithTemplate exportForm = new Gson().fromJson(json, ExportWithTemplate.class);
        ByteArrayInputStream templateInputStream = new ByteArrayInputStream(Resources.toByteArray(Resources.getResource("template/三管副学历生.xls")));

        exportForm.export(templateInputStream, new FileOutputStream("temp.xls"));
    }

    @Test
    void 导出Excel2003() throws IOException {
        String json = """
                {
                  "template": "三管副学历生",
                  "excelType": "xlsx",
                  "data": {
                    "list": [
                      {
                        "id": "200",
                        "idNumber":"640302196301122294",
                        "student": {
                          "id": "2",
                          "register": {
                            "name": "张三",
                            "idNumber": "640302196301122294",
                            "contactPhone": "12345678901",
                            "bundle": "A套餐",
                            "project": {
                              "id": 100,
                              "name": "船长"
                            },
                            "namePinyin": "ZhangSan"
                          },
                          "exam": {
                              "examType": "申考适任",
                              "dutyType": "船长"
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
                        "idNumber": "140221198509097336",
                        "student": {
                        "exam": {
                              "examType": "申考适任",
                              "dutyType": "船长"
                            },
                          "id": "200",
                          "register": {
                            "name": "李四",
                            "idNumber": "140221198509097336",
                            "namePinyin": "LiSi"
                          },
                          "recruitment": {
                            "id": 100,
                            "name": "船员培训-2023"
                          }
                        },
                        "index": "2"
                      },
                      {
                        "id": "200",
                         "idNumber":"640302196301122294",
                        "student": {
                        "exam": {
                              "examType": "申考适任",
                              "dutyType": "船长"
                            },
                          "id": "200",
                          "register": {
                            "name": "李四",
                            "idNumber": "140221198509097336",
                            "namePinyin": "LiSi"
                          },
                          "recruitment": {
                            "id": 100,
                            "name": "船员培训-2023"
                          }
                        },
                        "index": "2"
                      },
                      {
                        "id": "200",
                         "idNumber":"640302196301122294",
                        "student": {
                        "exam": {
                              "examType": "申考适任",
                              "dutyType": "船长"
                            },
                          "id": "200",
                          "register": {
                            "name": "李四",
                            "idNumber": "140221198509097336",
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
                  },
                  "config": {
                    "listName": "list",
                    "columns": [
                      {
                        "index": 1,
                        "name": "student.register.name"
                      },{
                        "index": 0,
                        "name": "student.register.idNumber"
                      }
                
                    ]
                  }
                }
                """;
        ExportWithTemplate exportForm = new Gson().fromJson(json, ExportWithTemplate.class);
        ByteArrayInputStream templateInputStream = new ByteArrayInputStream(Resources.toByteArray(Resources.getResource("template/三管副学历生.xlsx")));

        exportForm.export(templateInputStream, new FileOutputStream("temp.xlsx"));
    }

    @Test
    void 导出100() throws IOException {
        String json = """
                {
                  "template": "三管副学历生",
                  "excelType": "xls",
                  "data": {
                    "list": [{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}]
                  },
                  "config": {
                  }
                }
                """;
        ByteArrayInputStream templateInputStream = new ByteArrayInputStream(Resources.toByteArray(Resources.getResource("template/三管副学历生.xls")));

        ExportWithTemplate exportForm = new Gson().fromJson(json, ExportWithTemplate.class);
        exportForm.export(templateInputStream, new FileOutputStream("temp.xlsx"));
    }

    @Test
    void 通过单元格导出() throws IOException {
        String json = """
                {
                  "mergeRanges": [
                    {
                      "firstRow": 0,
                      "lastRow": 0,
                      "firstCol": 0,
                      "lastCol": 7
                    },
                    {
                      "firstRow": 1,
                      "lastRow": 2,
                      "firstCol": 0,
                      "lastCol": 0
                    },
                    {
                      "firstRow": 1,
                      "lastRow": 2,
                      "firstCol": 1,
                      "lastCol": 1
                    },
                    {
                      "firstRow": 1,
                      "lastRow": 1,
                      "firstCol": 2,
                      "lastCol": 3
                    },
                    {
                      "firstRow": 1,
                      "lastRow": 1,
                      "firstCol": 4,
                      "lastCol": 5
                    },
                    {
                      "firstRow": 1,
                      "lastRow": 2,
                      "firstCol": 6,
                      "lastCol": 6
                    },
                    {
                      "firstRow": 1,
                      "lastRow": 2,
                      "firstCol": 7,
                      "lastCol": 7
                    },
                    {
                      "firstRow": 4,
                      "lastRow": 4,
                      "firstCol": 0,
                      "lastCol": 1
                    }
                  ],
                  "rows": [
                    {
                      "cells": [
                        {
                          "content": "202306期船长",
                          "fontSize": 20,
                          "width": 20
                        },
                        {
                          "content": "123,222.12",
                          "fontSize": 12,
                          "width": 20,
                          "type": "number"
                        },
                        {
                          "content": "24.22%",
                          "fontSize": 12,
                          "width": 20,
                            "type": "percent"
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 50
                        }
                      ],
                      "height": 30
                    },
                    {
                      "height": 20,
                      "cells": [
                        {
                          "content": "准考证号码",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "姓名",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "理论成绩",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "评估成绩",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "总成绩",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "备注\\n（身份证号码）",
                          "fontSize": 12,
                          "width": 50
                        }
                      ]
                    },
                    {
                      "height": 60,
                      "cells": [
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "航海英语",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "航海学",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "船舶操纵与避碰",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "船舶结构与货运",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 50
                        }
                      ]
                    },
                    {
                      "cells": [
                        {
                          "content": null,
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "23%",
                          "fontSize": 12,
                          "width": 20,
                          "type":"percent"
                        },
                        {
                          "content": "123,344.23",
                          "fontSize": 12,
                          "width": 20,
                           "type":"number"
                        },
                        {
                          "content": "缺考",
                          "fontSize": 12,
                          "width": 20,
                          "font": {
                            "color": "RED"
                          }
                        },
                        {
                          "content": "及格",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "58.75",
                          "fontSize": 12,
                          "width": 20,
                          "type":"number"
                        },
                        {
                          "content": "不及格",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "110106198911010111",
                          "fontSize": 12,
                          "width": 50
                        }
                      ],
                      "height": 20
                    },
                    {
                      "cells": [
                        {
                          "content": "任课老师/通过率（%）",
                          "fontSize": 12,
                          "width": 20
                
                        },
                        {
                          "content": "",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "test/0%",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "test/0%",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "钟燕萍、test/100%",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "test/0%",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "一次性通过率0%",
                          "fontSize": 12,
                          "width": 20
                        },
                        {
                          "content": "一次性通过人数：0人",
                          "fontSize": 12,
                          "width": 50
                        }
                      ],
                      "height": 60
                    }
                  ]
                }
                """;
        ExportWithCells exportForm = new Gson().fromJson(json, ExportWithCells.class);
        exportForm.export(new FileOutputStream("temp.xlsx"));
    }

    @AfterEach
    public void tearDown() {
//        new java.io.File("temp.xlsx").delete();
//        new java.io.File("temp.xls").delete();
    }

}
