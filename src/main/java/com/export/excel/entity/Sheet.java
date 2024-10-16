package com.export.excel.entity;

import com.export.excel.entity.abs.Base;
import com.export.excel.entity.utils.StringUtils;

import java.awt.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class Sheet extends Base {

    /**
     * 宽度，设置宽度则支持固定宽度
     */
    private List<WidthStyle> widthStyles;

    private String name;

    public Sheet(String name) {
        // 校验sheet名称长度 (报表名称有些过长校验会导致导出被拦截，单机版未校验，暂不校验)
//        if (name.length() > 31) {
//            throw ExceptionHelper.newException(JpErrorCodeCst.SYSTEM_ERROR, "单个sheet名称长度不能超过31个字符，长度：" + name.length() + "，名称：" + name);
//        }
        this.name = name;
    }

    private List<Row> rows = new ArrayList<>();

    public String getName() {
        return name;
    }

    public List<Row> getRows() {
        return rows;
    }

    public List<WidthStyle> getWidthStyles() {
        return widthStyles;
    }

    public void setWidthStyles(List<WidthStyle> widthStyles) {
        this.widthStyles = widthStyles;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setRows(List<Row> rows) {
        this.rows = rows;
    }

    public void addRow(Row row) {
        // 默认排序
        if (row.getIndex() == -1) {
            row.setIndex(this.rows.size());
        }
        row.setParent(this);
        this.rows.add(row);
    }

    public void addRow(int index, Row row) {
        // 默认排序
        if (row.getIndex() == -1) {
            row.setIndex(this.rows.size());
        }
        row.setParent(this);
        this.rows.add(index, row);
    }

    public void setDefaultColumnWidth(short width) {
        for (Row row : getRows()) {
            List<Col> cols = row.getCols();
            for (Col col : cols) {
                if (col.style == null) {
                    col.style = ColStyle.DEFAULT;
                }
                if (col.style.width == 10) {
                    col.style.width = width;
                }
            }
        }
    }

    public void setDefaultRowHeightInPoints(short height) {
        for (Row row : getRows()) {
            if (row.style == null) {
                row.style = RowStyle.DEFAULT;
            }
            if (row.style.height == 18) {
                row.style.height = height;
            }
        }
    }

    /**
     * 初始化字体宽度 按列分组 index索引, width列宽
     *
     * @return Map<Integer, Integer> index索引, width列宽
     */
    public Map<Integer, Integer> getAutoSizeColumn() {
        // 初始化字体宽度 按列分组 index索引, width列宽
        Map<Integer, Integer> groupCol = new HashMap<>();
        for (Row row : this.rows) {
            for (Col col : row.getCols()) {
                int index = col.getIndex();
                if (col.style == null || col.style.font == null || StringUtils.isBlank(col.style.font.name)) {
                    continue;
                }
                short fontSize = col.style.font.size;
                String fontName = col.style.font.name;
                // 字体
                java.awt.Font font = new java.awt.Font(fontName, java.awt.Font.PLAIN, fontSize);
                if (!groupCol.containsKey(index)) {
                    int newWidth = getStringWidth(font, col.value);
                    if (col.colspan > 1) {
                        newWidth = newWidth / col.colspan;
                    }
                    for (int i = 0; i < col.colspan; i++) {
                        groupCol.put(index + i, newWidth);
                    }
                } else {
                    int newWidth = getStringWidth(font, col.value);
                    if (col.colspan > 1) {
                        newWidth = newWidth / col.colspan;
                    }
                    for (int i = 0; i < col.colspan; i++) {
                        if (groupCol.get(index + i) < newWidth) {
                            groupCol.put(index + i, newWidth);
                        }
                    }
                }
            }
        }
        return groupCol;
    }

    /**
     * 获取字体宽度
     */
    private int getStringWidth(java.awt.Font font, String str) {
        int rst = 0;
        if (StringUtils.isBlank(str)) {
            // 字符串为空跳过
            return rst;
        }
        char[] chars = str.toCharArray();
        for (char aChar : chars) {
            rst += getCharacterWidth(font, aChar);
        }
        // 最大宽度字节
        int max = 255;
        if (rst > max) {
            rst = max;
        }
        return rst;
    }

    /**
     * 获取字符宽度
     */
    private int getCharacterWidth(java.awt.Font font, char character) {
        Graphics graphics = new java.awt.image.BufferedImage(1, 1, java.awt.image.BufferedImage.TYPE_INT_RGB).getGraphics();
        graphics.setFont(font);
        FontMetrics fontMetrics = graphics.getFontMetrics(font);
        int width = fontMetrics.charWidth(character);
        graphics.dispose();
        return width / 3;
    }


}
