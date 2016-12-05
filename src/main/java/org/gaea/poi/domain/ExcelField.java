package org.gaea.poi.domain;

import java.io.Serializable;

/**
 * Created by iverson on 2016-6-4 16:20:57.
 */
public class ExcelField implements Serializable{
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
