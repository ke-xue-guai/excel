package com.export.excel.entity;

public class Font {

    /**
     * 默认字体样式
     */
    public static final Font DEFAULT = new Font();

    /**
     * 字号
     */
    public short size = 12;

    /**
     * 高度
     */
    public short height = 16;

    /**
     * 加粗
     */
    public boolean bold = false;

    /**
     * 下划线
     */
    public short underline = 0;

    /**
     * 字体名称
     */
    public String name = "宋体";

}
