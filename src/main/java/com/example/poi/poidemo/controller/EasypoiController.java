package com.example.poi.poidemo.controller;

import com.example.poi.poidemo.easypoi.base.EasypoiService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * @Description: easypoi使用
 * @Author xuyuc
 * @Date 2021/8/18
 */
@Controller
public class EasypoiController {

    @Resource
    EasypoiService easypoiService;

    // 表头 固定
    @GetMapping("/easypoiByClass")
    public void ExportFile1(HttpServletRequest request, HttpServletResponse response) {
        Workbook workbook = easypoiService.createWorkbookByClass();
        easypoiService.exportExcle(workbook, "文件名.xls", response);
    }

    // 表头 不固定
    @GetMapping("/easypoiByMap")
    public void ExportFile2(HttpServletRequest request, HttpServletResponse response) {
        Workbook workbook = easypoiService.createWorkbookByList();
        easypoiService.exportExcle(workbook, "文件名.xls", response);
    }
}
