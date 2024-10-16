package com.export.excel.entity.impl;

import com.alibaba.fastjson.JSON;
import com.export.excel.entity.*;
import com.export.excel.entity.Font;
import com.export.excel.entity.Row;
import com.export.excel.entity.Sheet;
import com.export.excel.entity.abs.IWorkBook;
import com.export.excel.entity.utils.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;


import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class HSSFWorkBook implements IWorkBook<HSSFSheet, HSSFRow, HSSFCell> {

    /**
     * 全局样式map，避免【The maximum number of Cell Styles was exceeded. You can define up to 64000 style in a .xlsx Workbook】
     */
    private Map<ColStyle, HSSFCellStyle> colStyleMap = new HashMap<>();

    /**
     * 全局字体样式map，避免创建过多的font
     */
    private Map<String, HSSFFont> fontStyleMap = new HashMap<>();

    private HSSFWorkbook workbook = new HSSFWorkbook();

    /**
     * 数据格式实例，用于转换数据格式
     */
    private HSSFDataFormat format = workbook.createDataFormat();

    private static final short BANG = 20;

    @Override
    public HSSFSheet create(Sheet sheet) {
        String name = sheet.getName();
        HSSFSheet hssfSheet = workbook.createSheet(name);
        List<WidthStyle> widthStyles = sheet.getWidthStyles();
        if (widthStyles == null) {
            Map<Integer, Integer> autoSizeColumn = sheet.getAutoSizeColumn();
            for (Integer index : autoSizeColumn.keySet()) {
                if (autoSizeColumn.get(index) == null || autoSizeColumn.get(index) == 0) {
                    hssfSheet.setColumnWidth(index, 10 * 256);
                } else {
                    hssfSheet.setColumnWidth(index, autoSizeColumn.get(index) * 256);
                }
            }
        } else {
            for (WidthStyle widthStyle : widthStyles) {
                hssfSheet.setColumnWidth(widthStyle.getIndex(), widthStyle.getWidth() * 256);
            }
        }
        return hssfSheet;
    }

    @Override
    public HSSFRow create(HSSFSheet sheet, Row row) {
        HSSFRow hssfRow = sheet.createRow(row.getIndex());
        hssfRow.setHeight((short) (row.style.height * BANG));
        return hssfRow;
    }

    @Override
    public HSSFCell create(HSSFSheet sheet, HSSFRow row, Col col) {
        int index = col.getIndex();
        HSSFCell cell = row.createCell(index);
        HSSFCellStyle style;

        if (col.style != null) {
            style = colStyleMap.get(col.style);
            if (style == null) {
                style = workbook.createCellStyle();
                //背景颜色
                this.setBackgroundColor(style, col.style.background);
                //数据格式
                this.setDataFormat(style, col.style.dataFormat);
                //水平
                style.setAlignment(HorizontalAlignment.forInt(col.style.align));
                //垂直
                style.setVerticalAlignment(VerticalAlignment.forInt(col.style.verticalAlign));
                // 边框
                this.setBorder(style, col.style.border);
                // 字体
                this.setFont(style, col.style.font);
                // 缓存
                colStyleMap.put(col.style, style);
            }
            //
            cell.setCellStyle(style);
        }
        String value = col.value;
        if (value != null) {
            if (col.style != null && StringUtils.isNotBlank(col.style.dataFormat)) {
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(Double.parseDouble(value));
            } else {
                cell.setCellType(CellType.STRING);
                cell.setCellValue(new HSSFRichTextString(value));
            }
        }
        // 合并单元格子
        if (col.rowspan > 1 || col.colspan > 1) {
            int rowIndex = col.getParent().getIndex();
            CellRangeAddress rangeAddress = new CellRangeAddress(rowIndex, rowIndex + col.rowspan - 1, index, index + col.colspan - 1);
            sheet.addMergedRegion(rangeAddress);
        }
        return cell;
    }

    public void setDataFormat(HSSFCellStyle cellStyle, String dataFormat) {
        // 数据格式为空的使用通用格式-0
        if (StringUtils.isBlank(dataFormat)) {
            cellStyle.setDataFormat((short) 0);
        } else {
            cellStyle.setDataFormat(format.getFormat(dataFormat));
        }
    }


    public void setBackgroundColor(HSSFCellStyle cellStyle, Short backgroundColor) {
        if (backgroundColor != -1) {
            //去掉边框
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            //设置背景颜色
            cellStyle.setFillForegroundColor(backgroundColor);
            cellStyle.setFillBackgroundColor(backgroundColor);
        }
    }

    public void setFont(HSSFCellStyle cellStyle, Font font) {
        String jsonString = JSON.toJSONString(font);
        if (fontStyleMap.get(jsonString) == null) {
            HSSFFont hssfFont = workbook.createFont();
            hssfFont.setFontHeightInPoints(font.size);
            hssfFont.setBold(font.bold);
            hssfFont.setFontName(font.name);
            hssfFont.setUnderline((byte) font.underline);
            fontStyleMap.put(jsonString, hssfFont);
        }
        HSSFFont hssfFont = fontStyleMap.get(jsonString);
        cellStyle.setFont(hssfFont);
    }

    public void setBorder(HSSFCellStyle cellStyle, Border border) {
        if (border == null) {
            return;
        }
        cellStyle.setBorderTop(BorderStyle.valueOf(border.getTop()));
        cellStyle.setBorderLeft(BorderStyle.valueOf(border.getLeft()));
        cellStyle.setBorderRight(BorderStyle.valueOf(border.getRight()));
        cellStyle.setBorderBottom(BorderStyle.valueOf(border.getBottom()));
    }

    @Override
    public ByteArrayOutputStream toByteArray() {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try {
            workbook.write(byteArrayOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return byteArrayOutputStream;
    }

}
