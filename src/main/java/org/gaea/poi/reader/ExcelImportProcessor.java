package org.gaea.poi.reader;

import org.gaea.exception.ProcessFailedException;
import org.gaea.exception.ValidationFailedException;

import java.io.InputStream;

/**
 * Created by iverson on 2016/6/8.
 */
public interface ExcelImportProcessor {
    void importDB(InputStream excelFileIS) throws ValidationFailedException, ProcessFailedException;
}
