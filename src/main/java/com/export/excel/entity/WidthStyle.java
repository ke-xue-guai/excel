package com.export.excel.entity;

import lombok.Getter;
import lombok.Setter;


@Getter
@Setter
public class WidthStyle {

    public static final WidthStyle DEFAULT = new WidthStyle();

    /**
     * 所在列
     */
    private Integer index;

    /**
     * 固定列宽的宽度
     */
    private Integer width;

}
