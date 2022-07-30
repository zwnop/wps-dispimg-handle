package com.zwnop.wps.dispimg.handle;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import lombok.Data;

@Data
@ExcelTarget("student")
public class Student {
    @Excel(name = "编号", orderNum = "0", type = 1)
    private String num;
    @Excel(name = "头像", orderNum = "1", type = 1)
    private String icon;
}
