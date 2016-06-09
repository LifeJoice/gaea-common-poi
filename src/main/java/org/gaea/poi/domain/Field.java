package org.gaea.poi.domain;

/**
 * XML定义的字段
 * Created by iverson on 2016-6-7 09:55:59.
 */
public class Field {
    private String name;
    private Integer columnIndex;
    private String readType;// 字段类型 : string(default) | date | time | datetime
    public static final String READ_TYPE_DATE = "date";
    public static final String READ_TYPE_TIME = "time";
    public static final String READ_TYPE_DATETIME = "datetime";

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

    public String getReadType() {
        return readType;
    }

    public void setReadType(String readType) {
        this.readType = readType;
    }
}
