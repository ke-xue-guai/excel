package com.export.excel.entity.abs;

public class Index extends Base {

    @Override
    public Index getParent() {
        return (Index) parent;
    }

    /**
     * 索引
     */
    private int index = -1;

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

}
