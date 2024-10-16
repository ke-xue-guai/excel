package com.export.excel.entity;

import com.export.excel.entity.abs.Index;

public class Col extends Index {

    /**
     * 值
     */
    public String value;

    /**
     * 跨行
     */
    public int rowspan = 1;

    /**
     * 跨列
     */
    public int colspan = 1;

    /**
     * 单元格样式
     */
    public ColStyle style;

    /**
     * 是否是空单元格
     */
    public Boolean isNull;

}
