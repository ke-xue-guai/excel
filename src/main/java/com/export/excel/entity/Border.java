package com.export.excel.entity;

public class Border {

    public static final short NONE = 0;
    public static final short THIN = 1;
    public static final short MEDIUM = 2;
    public static final short DASHED = 3;
    public static final short DOTTED = 4;
    public static final short THICK = 5;
    public static final short DOUBLE = 6;
    public static final short HAIR = 7;
    public static final short MEDIUM_DASHED = 8;
    public static final short DASH_DOT = 9;
    public static final short MEDIUM_DASH_DOT = 10;
    public static final short DASH_DOT_DOT = 11;
    public static final short MEDIUM_DASH_DOT_DOT = 12;
    public static final short SLANTED_DASH_DOT = 13;

    private short top = NONE;
    private short left = NONE;
    private short right = NONE;
    private short bottom = NONE;

    public Border() {

    }

    public Border(short top, short left, short right, short bottom) {
        this.top = top;
        this.left = left;
        this.right = right;
        this.bottom = bottom;
    }

    public short getTop() {
        return top;
    }

    public void setTop(short top) {
        this.top = top;
    }

    public short getLeft() {
        return left;
    }

    public void setLeft(short left) {
        this.left = left;
    }

    public short getRight() {
        return right;
    }

    public void setRight(short right) {
        this.right = right;
    }

    public short getBottom() {
        return bottom;
    }

    public void setBottom(short bottom) {
        this.bottom = bottom;
    }

}
