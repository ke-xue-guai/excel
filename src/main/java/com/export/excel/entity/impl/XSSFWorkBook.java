package com.export.excel.entity.impl;

import com.alibaba.fastjson.JSON;
import com.export.excel.entity.*;
import com.export.excel.entity.Font;
import com.export.excel.entity.Row;
import com.export.excel.entity.Sheet;
import com.export.excel.entity.abs.IWorkBook;
import com.export.excel.entity.utils.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class XSSFWorkBook implements IWorkBook<XSSFSheet, XSSFRow, XSSFCell> {

    private Map<ColStyle, XSSFCellStyle> colStyleMap = new HashMap<>();

    private Map<String, XSSFFont> fontStyleMap = new HashMap<>();

    private XSSFWorkbook workbook = new XSSFWorkbook();

    private XSSFDataFormat format = workbook.createDataFormat();

    private static final short BANG = 20;

    private int hasSetWidth = 0;

    @Override
    public XSSFSheet create(Sheet sheet) {
        String name = sheet.getName();
        XSSFSheet xssfSheet = workbook.createSheet(name);
        List<WidthStyle> widthStyles = sheet.getWidthStyles();
        if (widthStyles == null) {
            Map<Integer, Integer> autoSizeColumn = sheet.getAutoSizeColumn();
            for (Integer index : autoSizeColumn.keySet()) {
                if (autoSizeColumn.get(index) == null || autoSizeColumn.get(index) == 0) {
                    xssfSheet.setColumnWidth(index, 10 * 256);
                } else {
                    xssfSheet.setColumnWidth(index, autoSizeColumn.get(index) * 256);
                }
            }
        } else {
            for (WidthStyle widthStyle : widthStyles) {
                xssfSheet.setColumnWidth(widthStyle.getIndex(), widthStyle.getWidth() * 256);
            }
        }
        return xssfSheet;
    }

    @Override
    public XSSFRow create(XSSFSheet sheet, Row row) {
        XSSFRow xssfRow = sheet.createRow(row.getIndex());
        xssfRow.setHeight((short) (row.style.height * BANG));
        return xssfRow;
    }

    @Override
    public XSSFCell create(XSSFSheet sheet, XSSFRow row, Col col) {
        int index = col.getIndex();
        XSSFCell cell = row.createCell(index);
        XSSFCellStyle style;

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
                cell.setCellValue(new XSSFRichTextString(value));
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

    public void setDataFormat(XSSFCellStyle cellStyle, String dataFormat) {
        if (StringUtils.isNotBlank(dataFormat)) {
            cellStyle.setDataFormat(format.getFormat(dataFormat));
        }
    }


    public void setBackgroundColor(XSSFCellStyle cellStyle, Short backgroundColor) {
        if (backgroundColor != -1) {
            //去掉边框
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            //设置背景颜色
            cellStyle.setFillForegroundColor(backgroundColor);
            cellStyle.setFillBackgroundColor(backgroundColor);
        }
    }

    public void setFont(XSSFCellStyle cellStyle, Font font) {
        String jsonString = JSON.toJSONString(font);
        if (fontStyleMap.get(jsonString) == null) {
            XSSFFont xssfFont = workbook.createFont();
            xssfFont.setFontHeightInPoints(font.size);
            xssfFont.setBold(font.bold);
            xssfFont.setFontName(font.name);
            xssfFont.setUnderline((byte) font.underline);
            fontStyleMap.put(jsonString, xssfFont);
        }
        XSSFFont xssfFont = fontStyleMap.get(jsonString);
        cellStyle.setFont(xssfFont);
    }

    public void setBorder(XSSFCellStyle cellStyle, Border border) {
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
