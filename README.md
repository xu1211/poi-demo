
# apache.poi

用于操作office文档（word，excel）

## 1.导入依赖
```xml
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>x.x.x</version>
        </dependency>

        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>x.x.x</version>
        </dependency>
```
## 2.使用

### 写excel
1. 创建 Workbook 对象,将数据与单元格格式写入 Workbook
   >src/main/java/com/example/poi/poidemo/service/ExcelGeneratorService.java
2. 将Workbook 写出到文件
   >src/main/java/com/example/poi/poidemo/controller/ExcelController.java

# 其他poi
基本都是封装了apache.poi，使用简单，但对格式操作有局限

## easypoi
>http://doc.wupaas.com/docs/easypoi/

对apache.poi进行了封装，使用更加简单，但是输出格式有限制。复杂的格式无法实现

## Hutool-poi
>https://www.hutool.cn/docs/#/poi/概述

## GcExcel