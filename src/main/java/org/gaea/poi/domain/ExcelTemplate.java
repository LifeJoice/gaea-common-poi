package org.gaea.poi.domain;

import java.io.Serializable;
import java.util.List;

/**
 * Excel的模板对象。
 * 实现Serializable是为了可以缓存对象。
 * Created by iverson on 2016-11-2 17:46:33.
 */
public class ExcelTemplate implements Serializable{
    private String id;
    private List<Sheet> excelSheetList;
    private String fileName; // 文件名。以后扩展支持表达式语言，实现动态、个性化的文件名。例如：当前用户名+日期.xls

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public List<Sheet> getExcelSheetList() {
        return excelSheetList;
    }

    public void setExcelSheetList(List<Sheet> excelSheetList) {
        this.excelSheetList = excelSheetList;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }
}
