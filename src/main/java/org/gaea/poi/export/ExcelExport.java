package org.gaea.poi.export;

import org.gaea.exception.ProcessFailedException;
import org.gaea.exception.SysInitException;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.domain.Field;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Created by iverson on 2016/10/26.
 */
public interface ExcelExport {
    File export(String excelTemplateId, List<? extends Map> data, String fileDir, List<String> exportFieldNameList) throws SysInitException, ValidationFailedException, ProcessFailedException;

    File export(List<? extends Map> data, String title, Map<String, Field> fieldsMap, String fileName, String fileDir) throws ValidationFailedException, ProcessFailedException;
}
