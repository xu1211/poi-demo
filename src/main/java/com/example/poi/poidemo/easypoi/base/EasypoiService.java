package com.example.poi.poidemo.easypoi.base;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * @Description: easypoi使用
 * @Author xuyuc
 * @Date 2021/8/18
 */
@Service
public class EasypoiService {

    // 初始化 表数据
    static List<UserExportVo> exportList;

    public static void init() {
        exportList = Lists.newArrayList();
        for (int i = 0; i < 5; i++) {
            UserExportVo user = new UserExportVo();
            user.setUsername("user" + i);
            user.setPassword("pw" + i);
            exportList.add(user);
        }
    }

    /**
     * 使用情况：表头固定
     * 根据固定的表头创建：UserExportVo，表头字段添加 @Excel 注释
     *
     * @return
     */
    public Workbook createWorkbookByClass() {
        init();
        // 标题
        ExportParams exportParams = new ExportParams(null, "Sheet名", ExcelType.HSSF);

        // 表头
        Class userClass = UserExportVo.class;

        // Sheet = 标题 + 表头 + 表数据
        Map<String, Object> map = new HashMap<>();
        map.put("title", exportParams);
        map.put("entity", userClass);
        map.put("data", exportList);

        // Sheet集合
        List<Map<String, Object>> excelList = new ArrayList<>(1);
        excelList.add(map);

        // Sheet集合 生成 Workbook
        Workbook workbook = ExcelExportUtil.exportExcel(excelList, ExcelType.HSSF);
        return workbook;
    }

    /**
     * 使用情况 : 表头不固定
     * map 动态生成表头
     */
    public Workbook createWorkbookByList() {
        init();
        // 标题
        ExportParams exportParams = new ExportParams(null, "sheet名", ExcelType.HSSF);

        // 表头
        List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
        ExcelExportEntity excelentity = new ExcelExportEntity("姓名", "username");
        entity.add(excelentity);
        entity.add(new ExcelExportEntity("密码", "password"));

        // 生成workbook = 标题 + 表头 + 表数据
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, entity, exportList);
        return workbook;
    }

    /**
     * Workbook 转为 文件流
     *
     * @param workbook
     * @param fileName
     * @param response
     */
    public static void exportExcle(Workbook workbook, String fileName, HttpServletResponse response) {
        ServletOutputStream out = null;
        try {
            response.setContentType("application/vnd.ms-excel;chartset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            out = response.getOutputStream();
            workbook.write(out);
            out.flush();
            out.close();
        } catch (IOException e) {
            throw new RuntimeException("excel导出异常", e);
        }
    }
}
