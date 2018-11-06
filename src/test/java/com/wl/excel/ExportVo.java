package com.wl.excel;

import com.wl.excel.parser.Excel;
import com.wl.excel.parser.ExcelColumn;

/**
 * 测试vo
 *
 * @author wl
 * @date 2017/11/8.
 */
@Excel(titleRowIndex = 2, dataRowIndex = 4)
public class ExportVo {

    @ExcelColumn(1)
    private String userName;

    @ExcelColumn(2)
    private int age;

    @ExcelColumn(3)
    private String sex;

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
