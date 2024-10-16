package com.export.excel.entity;

import com.export.excel.entity.abs.Base;
import com.export.excel.entity.utils.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

public class Excel extends Base {

    private List<Sheet> sheets = new ArrayList<>();

    public List<Sheet> getSheets() {
        return sheets;
    }

    /**
     * 添加一个sheet
     */
    public void addSheet(Sheet sheet) {
        // 不能有重复的sheet名称
        Optional<Sheet> optional = sheets.stream().filter(it -> StringUtils.equals(it.getName(), sheet.getName())).findAny();
        if (optional.isPresent()) {
            throw new RuntimeException("Excel sheet's name is exist");
        }
        sheet.setParent(this);
        sheets.add(sheet);
    }

    public int getNumberOfSheets() {
        return sheets.size();
    }
}
