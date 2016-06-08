package org.gaea.poi.domain;

/**
 * Created by iverson on 2016-6-4 16:20:57.
 */
public class ExcelField {
    private String name;
    private Integer columnIndex;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(Integer columnIndex) {
        this.columnIndex = columnIndex;
    }
}
