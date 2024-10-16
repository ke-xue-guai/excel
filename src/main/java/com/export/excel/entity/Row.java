package com.export.excel.entity;

import com.export.excel.entity.abs.Index;
import com.export.excel.entity.utils.StringUtils;

import java.util.ArrayList;
import java.util.List;

public class Row extends Index {

    /**
     * 行样式
     */
    public RowStyle style = new RowStyle();

    public Row() {

    }

    private List<Col> cols = new ArrayList<>();

    public List<Col> getCols() {
        return this.cols;
    }

    public void addCol(Col col) {
        // 默认排序
        if (col.getIndex() == -1) {
            col.setIndex(this.cols.size());
        }
        col.setParent(this);
        this.cols.add(col);
    }

    public Integer getLastColNum() {
        return this.getCols().size();
    }

    /**
     * 该行单元格值是否全为空或者全为空字符串
     *
     * @return
     */
    public boolean isBlank() {
        for (Col col : this.cols) {
            String value = col.value;
            if (StringUtils.isBlank(value)){
                continue;
            }
            if (StringUtils.isNotBlank(value.replace(" ", ""))) {
                return false;
            }
        }
        return true;
    }
}
