package org.gaea.poi;

import org.gaea.exception.ValidationFailedException;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * Created by iverson on 2016-6-3 17:58:06
 */
public interface ExcelReader {
    org.gaea.poi.domain.Workbook getWorkbook(InputStream fileIS) throws ValidationFailedException;

    List<Map<String, String>> getData(InputStream fileIS) throws ValidationFailedException;

    <T> List<T> getData(InputStream fileIS, Class<T> beanClass) throws ValidationFailedException;
}
