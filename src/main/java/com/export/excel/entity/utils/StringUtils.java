package com.export.excel.entity.utils;

import java.util.Objects;

public class StringUtils {

    public static boolean isBlank(String para) {
        return para == null || para.equals("");
    }

    public static boolean isNotBlank(String para) {
        return !isBlank(para);
    }

    public static boolean equals(String a, Object b) {
        return Objects.equals(a, b);
    }
}
