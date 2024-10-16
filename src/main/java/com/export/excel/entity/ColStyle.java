package com.export.excel.entity;


public class ColStyle {

    public static final ColStyle DEFAULT = new ColStyle();

    /**
     * 代表背景灰度25的数字
     */
    public static final short GREY_25_PERCENT = 22;
    /**
     * 代表背景颜色white的数字
     */
    public static final short EXCEL_WHITE = 9;

    /**
     * 垂直对其方式
     */
    public static final int TOP = 0;
    public static final int MIDDLE = 1;
    public static final int BOTTOM = 2;
    /**
     * 水平对齐方式
     */
    public static final int LEFT = 1;
    public static final int CENTER = 2;
    public static final int RIGHT = 3;

    /**
     * 数据类型
     */
    public static final String TYPE_INT = "####";
    public static final String TYPE_DOUBLE_PRECISION = "###0.00";
    public static final String TYPE_MILLENNIAL = "#,###";
    public static final String TYPE_DOUBLE_PRECISION_MILLENNIAL = "#,##0.00";

    public short width = 10;

    /**
     * 水平对其方式
     */
    public int align = LEFT;

    /**
     * 垂直对其方式
     */
    public int verticalAlign = MIDDLE;

    /**
     * 背景颜色
     */
    public short background = -1;

    /**
     * 数据格式
     */
    public String dataFormat;

    /**
     * 字体
     */
    public Font font = Font.DEFAULT;

    /**
     * 边框
     */
    public Border border;

    public ColStyle copy() {
        ColStyle rst = new ColStyle();
        rst.width = width;
        rst.align = align;
        rst.verticalAlign = verticalAlign;
        rst.background = background;
        rst.dataFormat = dataFormat;
        rst.font = font;
        rst.border = border;
        return rst;
    }

}
