package com.export.excel;

import com.export.excel.entity.*;
import com.export.excel.entity.Font;
import com.export.excel.entity.Row;
import com.export.excel.entity.Sheet;
import com.export.excel.entity.abs.IWorkBook;
import com.export.excel.entity.impl.HSSFWorkBook;
import com.export.excel.entity.impl.XSSFWorkBook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.List;
import java.util.stream.Collectors;


public class ExcelExportUtils<S, R, C> {

    private IWorkBook<S, R, C> workbook;

    /**
     * 导出xls
     */
    public static ByteArrayOutputStream toXls(Excel excel) {
        ExcelExportUtils<HSSFSheet, HSSFRow, HSSFCell> excelExportUtils = new ExcelExportUtils<>();
        excelExportUtils.workbook = new HSSFWorkBook();
        excelExportUtils.export(excel);
        return excelExportUtils.workbook.toByteArray();
    }

    /**
     * 导出xlsx
     */
    public static ByteArrayOutputStream toXlsx(Excel excel) {
        ExcelExportUtils<XSSFSheet, XSSFRow, XSSFCell> excelExportUtils = new ExcelExportUtils<>();
        excelExportUtils.workbook = new XSSFWorkBook();
        excelExportUtils.export(excel);
        return excelExportUtils.workbook.toByteArray();
    }

    private void export(Excel excel) {
        for (Sheet sheet : excel.getSheets()) {
            sheet = this.mergeStyles(sheet);
            S hssfSheet = workbook.create(sheet);
            writeRow(hssfSheet, sheet.getRows());
        }
    }

    private void writeRow(S sheet, List<Row> rows) {
        for (int i = 0; i < rows.size(); i++) {
            Row row = rows.get(i);
            R hssfRow = workbook.create(sheet, row);
            writeCol(sheet, hssfRow, row.getCols());
        }
    }

    private void writeCol(S sheet, R row, List<Col> cols) {
        for (int i = 0; i < cols.size(); i++) {
            Col col = cols.get(i);
            workbook.create(sheet, row, col);
        }
    }

    /**
     * 解析返回Excel
     */
    public static Excel toExcel(InputStream inputStream) {
        Workbook workbook = null;
        try {
            // 用来解决压缩炸弹，设置压缩文件与扩展数据大小的最大值
            ZipSecureFile.setMinInflateRatio(0.001);
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        Excel excel = new Excel();
        // 判断整个Excel是不是空
        if (workbook != null) {
            // 获取一共有多少个sheet页面  getPhysicalNumberOfRows 不包含空单元格 getLastCellNum包含空单元格
            int sheetSize = workbook.getNumberOfSheets();
            for (int i = 0; i < sheetSize; i++) {
                String sheetName = workbook.getSheetAt(i).getSheetName();
                Sheet sheet = new Sheet(sheetName);
                int rowSize = workbook.getSheetAt(i).getLastRowNum();
                for (int j = 0; j <= rowSize; j++) {
                    Row row = new Row();
                    if (workbook.getSheetAt(i).getRow(j) == null) {
                        sheet.addRow(row);
                        continue;
                    }
                    for (int k = 0; k < workbook.getSheetAt(i).getRow(j).getLastCellNum(); k++) {
                        Cell cell = workbook.getSheetAt(i).getRow(j).getCell(k);
                        if (cell == null) {
                            Col col = new Col();
                            col.isNull = true;
                            row.addCol(col);
                            continue;
                        }
                        Col col = new Col();
                        col.setIndex(k);
                        CellStyle cellStyle = workbook.getSheetAt(i).getRow(j).getCell(k).getCellStyle();

                        // 通过单元格类型填写value
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case NUMERIC:
                                col.value = toStr(cell.getNumericCellValue());
                                break;
                            case STRING:
                                col.value = cell.getStringCellValue();
                                break;
                            case BOOLEAN:
                                col.value = toStr(cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                // 创建一个公式求值器
                                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                                // 设置求值器的上下文为当前单元格所在的工作表
                                evaluator.evaluateFormulaCell(cell);
                                // 从求值器中获取计算结果
                                CellValue cellValue = evaluator.evaluate(cell);
                                // 根据单元格类型，获取相应的值
                                if (cellValue.getCellType() == CellType.STRING) {
                                    String value = cellValue.getStringValue();
                                    col.value = value;
                                    break;
                                    // 处理字符串值
                                } else if (cellValue.getCellType() == CellType.NUMERIC) {
                                    double value = cellValue.getNumberValue();
                                    col.value = toStr(value);
                                    break;
                                    // 处理数值值
                                }
                                break;
                            default:
                                col.value = null;
                        }
                        col.style = new ColStyle();
                        col.style.align = cellStyle.getAlignment().getCode();
                        col.style.verticalAlign = cellStyle.getVerticalAlignment().getCode();
                        Border border = new Border();
                        border.setTop(cellStyle.getBorderTop().getCode());
                        border.setRight(cellStyle.getBorderRight().getCode());
                        border.setLeft(cellStyle.getBorderLeft().getCode());
                        border.setBottom(cellStyle.getBorderBottom().getCode());
                        col.style.border = border;
                        setFont(workbook, col, cellStyle);
                        String colorByCell = getColorByCell(cellStyle);
                        if (colorByCell != null) {
                            col.style.background = ColStyle.GREY_25_PERCENT;
                        } else {
                            col.style.background = ColStyle.EXCEL_WHITE;
                        }
                        System.out.println(colorByCell + col.value);
                        col.style.width = (short) (workbook.getSheetAt(i).getColumnWidth(j) / 256);
                        row.addCol(col);
                    }
                    sheet.addRow(row);
                }
                List<CellRangeAddress> mergedRegions = workbook.getSheetAt(i).getMergedRegions();
                for (CellRangeAddress mergedRegion : mergedRegions) {
                    Col colStart = sheet.getRows().get(mergedRegion.getFirstRow()).getCols().get(mergedRegion.getFirstColumn());
                    colStart.rowspan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                    colStart.colspan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
                    for (int j = mergedRegion.getFirstRow(); j <= mergedRegion.getLastRow(); j++) {
                        for (int k = mergedRegion.getFirstColumn(); k <= mergedRegion.getLastColumn(); k++) {
                            if (j == mergedRegion.getFirstRow() && k == mergedRegion.getFirstColumn()) {
                                continue;
                            }
                            if (sheet.getRows().get(j).getCols().size() <= k) {
                                break;
                            }
                            sheet.getRows().get(j).getCols().get(k).rowspan = -1;
                            sheet.getRows().get(j).getCols().get(k).colspan = -1;
                            sheet.getRows().get(j).getCols().get(k).value = "";
                        }
                    }
                }
                excel.addSheet(sheet);
            }
        }
        return excel;
    }

    /**
     * 设置字体
     */
    private static void setFont(Workbook workbook, Col col, CellStyle cellStyle) {
        short fontHeightInPoints;
        if (cellStyle instanceof HSSFCellStyle) {
            fontHeightInPoints = ((HSSFCellStyle) cellStyle).getFont(workbook).getFontHeightInPoints();
        } else if (cellStyle instanceof XSSFCellStyle) {
            fontHeightInPoints = ((XSSFCellStyle) cellStyle).getFont().getFontHeightInPoints();
        } else {
            fontHeightInPoints = Font.DEFAULT.size;
        }
        Font font = new Font();
        font.size = fontHeightInPoints;
        col.style.font = font;
    }

    /**
     * 转字符串
     *
     * @param object object
     * @return return
     */
    public static String toStr(Object object) {
        if (object == null) {
            return "";
        }
        if (object instanceof String) {
            return (String) object;
        }
        if (object instanceof BigDecimal) {
            return ((BigDecimal) object).stripTrailingZeros().toPlainString();
        }
        if (object instanceof Integer) {
            return BigDecimal.valueOf((Integer) object).stripTrailingZeros().toPlainString();
        }
        if (object instanceof Double) {
            return BigDecimal.valueOf((Double) object).stripTrailingZeros().toPlainString();
        }
        return object.toString();
    }

    /**
     * 合并单元格样式，通过给每一个单元格附加样式
     *
     * @param sheet sheet
     * @return 合并的Sheet
     */
    public Sheet mergeStyles(Sheet sheet) {
        List<WidthStyle> widthStyles = sheet.getWidthStyles();
        sheet = this.fillCol(sheet);
        List<Row> rows = sheet.getRows();
        int rowSize = rows.size();
        for (int i = 0; i < rowSize; i++) {
            int colSize = rows.get(i).getCols().size();
            for (int k = 0; k < colSize; k++) {
                Col col = rows.get(i).getCols().get(k);
                int currentRow = rows.get(i).getIndex();
                int currentCol = col.getIndex();
                int endRow = col.rowspan + currentRow;
                int endCol = col.colspan + currentCol;
                if (col.rowspan > 1 || col.colspan > 1) {
                    ColStyle style = rows.get(i).getCols().get(k).style;
                    for (int l = currentRow; l < endRow; l++) {
                        for (int m = currentCol; m < endCol; m++) {
                            rows.get(l).getCols().get(m).style = style;
                        }
                    }
                }
            }
        }
        sheet.setWidthStyles(widthStyles);
        return sheet;
    }

    /**
     * 填充单元格
     *
     * @param sheet sheet
     * @return 填充后的单元格
     */
    private Sheet fillCol(Sheet sheet) {
        Sheet resultSheet = new Sheet(sheet.getName());
        int locs = 0;
        int maxRow = 0;
        //获取最大列的下标
        for (Row row : sheet.getRows()) {
            for (Col col : row.getCols()) {
                locs = Math.max(locs, col.getIndex());
            }
            maxRow = Math.max(maxRow, row.getIndex());
        }
        // 对列中单元格进行填充
        for (int i = 0; i <= maxRow; i++) {
            Row row = sheet.getRows().get(i);
            // 填补缺少的行
            while (i < row.getIndex()) {
                Row resultRow = new Row();
                for (int j = 0; j < locs; j++) {
                    resultRow.addCol(new Col());
                }
                resultRow.setIndex(i);
                sheet.addRow(i, resultRow);
                resultSheet.addRow(i, resultRow);
                i++;
            }

            Row resultRow = new Row();
            resultRow.style = row.style;
            int k = 0;
            // k记录读取旧的的列号,j记录当前插入resultRow的位置index
            for (int j = 0; j <= locs; j++) {
                try {
                    Col col = row.getCols().get(k);
                    for (; j < col.getIndex(); j++) {
                        Col newCol = new Col();
                        newCol.setIndex(j);
                        resultRow.addCol(newCol);
                    }
                    resultRow.addCol(col);
                } catch (Exception e) {
                    Col newCol = new Col();
                    newCol.setIndex(j);
                    resultRow.addCol(newCol);
                }
                k++;
            }
            //通过index对数据覆盖,覆盖前需要判断单元格是否存在，如果存在跨行或跨列则不需要覆盖值
            for (int j = 0; j < row.getCols().size(); j++) {
                Col col = row.getCols().get(j);
                int index = col.getIndex();
                resultRow.getCols().set(index, col);
            }
            //去重，目的是去掉下标一样的，防止合并冲突
            List<Col> collect = resultRow.getCols().stream().distinct().collect(Collectors.toList());
            Row collectRow = new Row();
            collectRow.style = resultRow.style;
            for (Col col : collect) {
                collectRow.addCol(col);
            }
            //将col补齐到最大列
            for (int j = 0; j < resultRow.getCols().size() - collect.size(); j++) {
                Col col = new Col();
                collectRow.addCol(col);
            }
            resultSheet.addRow(collectRow);
        }
        return resultSheet;
    }


    /**
     * 获取三原色（支持xlsx和xls）
     *
     * @param style 单元格样式
     * @return
     */
    private static String getColorByCell(CellStyle style) {
        if (style.getFillForegroundColorColor() instanceof XSSFColor) {
            XSSFColor color = (XSSFColor) style.getFillForegroundColorColor();
            if (color != null) {
                short indexed = color.getIndexed();
                if (color.isRGB()) {
                    byte[] bytes = color.getRGB();
                    if (bytes != null && bytes.length == 3) {
                        int sum = 0;
                        StringBuilder sb = new StringBuilder();
                        sb.append("rgb");
                        sb.append("(");
                        for (int i = 0; i < bytes.length; i++) {
                            byte b = bytes[i];
                            int temp;
                            if (b < 0) {
                                temp = 256 + (int) b;
                            } else {
                                temp = b;
                            }
                            sum += temp;
                            sb.append(temp);
                            if (i != bytes.length - 1) {
                                sb.append(",");
                            }
                        }
                        sb.append(")");
                        if (sum > 0) {
                            return sb.toString();
                        }
                    }
                }
                if (ColStyle.GREY_25_PERCENT == indexed) {
                    return String.valueOf(indexed);
                }
            }
        }
        if (style.getFillForegroundColorColor() instanceof HSSFColor) {
            HSSFColor color = (HSSFColor) style.getFillForegroundColorColor();
            if (color != null) {
                short[] triplet = color.getTriplet();
                // 255+255+255为 白色
                if (triplet[0] + triplet[1] + triplet[2] > 0 && triplet[0] + triplet[1] + triplet[2] != 255 * 3) {
                    return "(" + triplet[0] + "," + triplet[1] + "," + triplet[2] + ")";
                }
            }
        }
        return null;
    }
}
