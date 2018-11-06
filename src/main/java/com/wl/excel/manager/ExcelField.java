package com.wl.excel.manager;

/**
 * excel字段
 * @author wl
 * @date 2017/11/8
 */
public class ExcelField {

    /**field名称*/
    private String fieldName;

    /**field在excel中的位置索引*/
    private int columnIndex;

    /**对应的excel中名称*/
    private String columnName;

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(int columnIndex) {
        this.columnIndex = columnIndex;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }
}
