package com.example.poi.poidemo.controller;

import com.example.poi.poidemo.service.ExcelGeneratorService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

/**
 * @Description: excel操作
 * @Author xuyuc
 * @Date 2021/6/28
 */
@Controller
public class ExcelController {

    @Resource
    ExcelGeneratorService excelGeneratorService;

    /**
     * workbook 写出到 流
     * （生成本地文件）
     */
    @GetMapping("/getFile")
    public void ExportFile(HttpServletRequest request, HttpServletResponse response) {
        HSSFWorkbook workbook = excelGeneratorService.createWorkbook();
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream("D:/excel导出.xls");
            workbook.write(fos);
        } catch (Exception e) {
            throw new RuntimeException("excel导出异常", e);
        } finally {
            try {
                if (fos != null) {
                    fos.close();
                }
            } catch (Exception e) {
                throw new RuntimeException("excel导出异常", e);
            }
        }
    }


    /**
     * workbook 写出到 response
     * （用于http下载excel）
     */
    @GetMapping("/getExcel")
    public void ExportExcel(HttpServletRequest request, HttpServletResponse response) {
        HSSFWorkbook workbook = excelGeneratorService.createWorkbook();
        OutputStream os = null;
        try {
            String fileName = new String("excel导出.xls".getBytes(), StandardCharsets.UTF_8);
            response.reset();
            response.addHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.flushBuffer();
            os = response.getOutputStream();
            workbook.write(os);
            os.flush();
        } catch (Exception e) {
            throw new RuntimeException("excel导出异常", e);
        } finally {
            try {
                if (os != null) {
                    os.close();
                }
            } catch (Exception e) {
                throw new RuntimeException("excel导出异常", e);
            }
        }
    }
}
