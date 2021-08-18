
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

---

# 其他poi
基本都是封装了apache.poi，使用简单，但对格式操作有局限

## easypoi
官方文档：
>http://doc.wupaas.com/docs/easypoi/


支持的场景格式：
1. 表头+表数据
   1.  表头固定\
   创建一个pojo 对应表头，使用@Excel等注解字段 
   2. 表头不固定\
      List<ExcelExportEntity> 表头
      List<Map<String, Object>> 表数据

2. 模板导出\
指定的单元格 填入 指定的变量


对apache.poi进行了封装，使用更加简单，但是输出格式有限制。复杂的格式无法实现

## 1.导入依赖
```xml
<dependency>
   <groupId>cn.afterturn</groupId>
   <artifactId>easypoi-spring-boot-starter</artifactId>
   <version>4.2.0</version>
</dependency>
```
## 2.使用

### 写excel
- 场景1 ： 表头固定
- 场景2 ： 表头不固定
- 场景3 ： 多sheet



## Hutool-poi
>https://www.hutool.cn/docs/#/poi/概述

## GcExcel