package com.export.excel.entity.abs;

import com.export.excel.entity.Col;
import com.export.excel.entity.Row;
import com.export.excel.entity.Sheet;

import java.io.ByteArrayOutputStream;

public interface IWorkBook<S, R, C> {

    S create(Sheet sheet);

    R create(S sheet, Row row);

    C create(S sheet, R row, Col col);

    ByteArrayOutputStream toByteArray();

}
