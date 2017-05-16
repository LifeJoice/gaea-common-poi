package org.gaea.poi.service;

import org.apache.poi.ss.usermodel.*;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.domain.*;
import org.gaea.poi.domain.Workbook;

import java.io.InputStream;
import java.util.Map;

/**
 * Created by iverson on 2017/5/10.
 */
public interface ExcelDefineService {
    Workbook getWorkbookDefine(InputStream excelIS, org.apache.poi.ss.usermodel.Workbook apacheWorkbook) throws ValidationFailedException;

    ExcelSheet getSheetDefine(Cell cell) throws ValidationFailedException;

    Block blockParse(org.gaea.poi.domain.Workbook workbook, ExcelBlock excelBlock);

    String getCellComment(Cell cell);

    Map<Integer, Field> getFieldsDefine(Row row, Map<String, Field> fieldDefMap) throws ValidationFailedException;

    ExcelField getField(Cell cell) throws ValidationFailedException;

    Field toField(ExcelField excelField);

    void combineField(Block block, ExcelField excelField);
}
