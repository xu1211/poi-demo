package com.example.poi.poidemo.service;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import sun.misc.BASE64Decoder;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import java.io.IOException;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Description: excel档案处理服务
 * @Author xuyuc
 * @Date 2021/6/24
 */
@Service
public class ExcelGeneratorService {


    public HSSFWorkbook createWorkbook() {
        /**
         * 创建 excel对象
         */
        HSSFWorkbook workbook = new HSSFWorkbook();//excel文件对象
        /**
         * 创建 excel表格对象
         */
        HSSFSheet sheet = workbook.createSheet("Sheet1");//excel工作表对象


        // 设置列宽
        sheet.setColumnWidth((short) 0, (short) 10 * 256);
        sheet.setColumnWidth((short) 1, (short) 10 * 256);
        sheet.setColumnWidth((short) 2, (short) 10 * 256);

        /**
         * 单元格 格式准备
         */
        // 宋体 18 加粗 水平居左 垂直居中
        HSSFCellStyle style1 = workbook.createCellStyle();
        Font font1 = workbook.createFont();
        font1.setFontHeightInPoints((short) 18);
        font1.setFontName("宋体");
        font1.setBold(true);
        style1.setAlignment(HorizontalAlignment.LEFT);
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        style1.setFont(font1);

        // 宋体 16 加粗 水平居左 垂直居中 背景色：LIGHT_TURQUOISE
        HSSFCellStyle style2 = workbook.createCellStyle();
        Font font2 = workbook.createFont();
        font2.setFontHeightInPoints((short) 16);
        font2.setFontName("宋体");
        font2.setBold(true);
        style2.setAlignment(HorizontalAlignment.LEFT);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setFont(font2);
        //上下左右边框：THIN
        style2.setBorderTop(BorderStyle.THIN);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);


        //插入图片
        String base64Str = "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/";
        BASE64Decoder decoder = new BASE64Decoder();
        byte[] decoderBytes = new byte[0];
        try {
            decoderBytes = decoder.decodeBuffer(base64Str);
        } catch (IOException e) {
            e.printStackTrace();
        }
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = new HSSFClientAnchor(800, 200, 0, 0, (short) 0, 0, (short) 2, 7);
        patriarch.createPicture(anchor, workbook.addPicture(decoderBytes, HSSFWorkbook.PICTURE_TYPE_JPEG));

        //写标题
        HSSFCell cell03 = sheet.createRow(0).createCell(3);
        cell03.setCellStyle(style1);
        cell03.setCellValue("标签1");

        HSSFCell cell13 = sheet.createRow(2).createCell(3);
        cell13.setCellStyle(style1);
        cell13.setCellValue("标签2");

        //单元格合并处理
        setMerged(sheet, "A1:C8");
        setMerged(sheet, "D1:L2");
        setMerged(sheet, "D3:L3");
        setMerged(sheet, "D4:L4");
        setMerged(sheet, "D5:L5");
        setMerged(sheet, "D6:L6");
        setMerged(sheet, "D7:L8");


        /**
         *  循环处理数据
         */
        Integer count = 8;//渲染行计数

        //标题 -行对象
        HSSFCell cellNavigationbar = null;
        for (int naviCount = 0; naviCount < 2; naviCount++) {
            cellNavigationbar = sheet.createRow(count).createCell(0);
            cellNavigationbar.setCellStyle(style2);
            cellNavigationbar.setCellValue("标题" + naviCount);
            count += 1;
            setMerged(sheet, "A" + count.toString() + ":L" + count.toString());

            for (int headRow = 0; headRow < 3; headRow++) {
                //表头 -行对象
                HSSFRow tableConfRow = sheet.createRow(count);
                for (int headCell = 0; headCell < 12; headCell++) {
                    //表头 -单元格对象
                    HSSFCell cellTable = tableConfRow.createCell(headCell);
                    cellTable.setCellStyle(style1);
                    cellTable.setCellValue("Head" + headCell);
                }
                count += 1;

                for (int dataRow = 0; dataRow < 4; dataRow++) {
                    //数据 -行对象
                    HSSFRow fieldDateRow = sheet.createRow(count);
                    for (int dataCell = 0; dataCell < 12; dataCell++) {
                        //数据 -单元格
                        HSSFCell fieldDateCell = fieldDateRow.createCell(dataCell);
                        fieldDateCell.setCellValue("data" + dataRow + dataCell);
                    }
                    count += 1;
                }
            }
        }
        return workbook;
    }

    /**
     * 合并单元格，设置单元格格式
     *
     * @param sheet excel工作表对象
     * @param ref   单元格“A1:B1”
     */
    private void setMerged(Sheet sheet, String ref) {
        CellRangeAddress region1 = CellRangeAddress.valueOf(ref);
        sheet.addMergedRegion(region1);
        setBorderStyle(sheet, region1);
    }


    /**
     * 设置合并单元格边框 - 线条
     */
    private void setBorderStyle(Sheet sheet, CellRangeAddress region) {
        // 合并单元格左边框样式
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        // 合并单元格上边框样式
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        // 合并单元格右边框样式
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
        // 合并单元格下边框样式
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
    }
}
