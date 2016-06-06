package org.gaea.poi;

import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.domain.GaeaPoiResultGroup;

import java.io.InputStream;

/**
 * Created by iverson on 2016-6-3 17:58:06
 */
public interface ExcelReader {
    <T> GaeaPoiResultGroup<T> getData(InputStream fileIS, Class<T> beanClass) throws ValidationFailedException;
}
