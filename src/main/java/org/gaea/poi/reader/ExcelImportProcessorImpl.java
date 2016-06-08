package org.gaea.poi.reader;

import org.gaea.exception.ProcessFailedException;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.ExcelReader;
import org.gaea.poi.domain.Block;
import org.gaea.poi.domain.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.InputStream;

/**
 * Created by iverson on 2016-6-8 11:00:03.
 */
@Component
public class ExcelImportProcessorImpl implements ExcelImportProcessor {
    private final Logger logger = LoggerFactory.getLogger(ExcelImportProcessorImpl.class);
    @Autowired
    private ImportDBProcessor importDBProcessor;
    @Autowired
    private ExcelReader excelReader;

    public void importDB(InputStream excelFileIS) throws ValidationFailedException, ProcessFailedException {
        Workbook workbook = excelReader.getWorkbook(excelFileIS);
        if(workbook==null){
            throw new ProcessFailedException("解析excel失败，获得空的解析结果。");
        }
        if(workbook.getBlockList()!=null){
            for(Block block:workbook.getBlockList()){
                importDBProcessor.executeImport(block);
            }
        }
    }
}
