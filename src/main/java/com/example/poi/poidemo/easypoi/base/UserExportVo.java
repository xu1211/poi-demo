package com.example.poi.poidemo.easypoi.base;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * @Description: ecxel表头
 * @Author xuyuc
 * @Date 2021/8/18
 */

@Data
public class UserExportVo {
    /**
     * 姓名
     */
    @Excel(name = "用户名", orderNum = "1", width = 40)
    private String username;
    /**
     * 密码
     */
    @Excel(name = "密码", orderNum = "2", width = 40)
    private String password;

}
