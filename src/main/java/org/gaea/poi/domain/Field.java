package org.gaea.poi.domain;

import java.io.Serializable;

/**
 * XML定义的字段
 * <p>实现Serializable是为了可以缓存对象。</p>
 * Created by iverson on 2016-6-7 09:55:59.
 */
public class Field implements Serializable{
    private String name;
    private Integer columnIndex;
    private String dataType;// 字段类型 : string(default) | date | time | datetime
    public static final String DATA_TYPE_STRING = "string";
    public static final String DATA_TYPE_DATE = "date";
    public static final String DATA_TYPE_TIME = "time";
    public static final String DATA_TYPE_DATETIME = "datetime";
    public static final String DATA_TYPE_NUMBER = "number";
    public static final String DATA_TYPE_INTEGER = "integer";
    public static final String DATA_TYPE_DOUBLE = "double";
    private String datetimeFormat; // 日期类型列的格式, 正常datatype=date,time之类的才有值. 参考Java SimpleDateFormat.
    private String title; // column的title text.
    private String cellValueTransferBy = "default"; // 对cell的值的处理。默认就是object.toString
    public static final String TRANSFER_BY_DEFAULT = "default"; // 默认默认1 就是object.toString
    public static final String TRANSFER_BY_DS_VALUE = "ds_value"; // 默认按数据集的value（Item.value）转换
    public static final String TRANSFER_BY_DS_TEXT = "ds_text"; // 默认按数据集的text（Item.text）转换
    private String width = "10"; // 单元格的宽度
    private String dataSetId; // 数据集的id。用来做值的转换。
    private String dbColumnName; // 数据库字段名. 基于SQL导出用。
    private String titleComment; // 标题行的注解
    private boolean visible = true; // 列是否可见

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

    public String getDataType() {
        return dataType;
    }

    public void setDataType(String dataType) {
        this.dataType = dataType;
    }

    public String getDatetimeFormat() {
        return datetimeFormat;
    }

    public void setDatetimeFormat(String datetimeFormat) {
        this.datetimeFormat = datetimeFormat;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getCellValueTransferBy() {
        return cellValueTransferBy;
    }

    public void setCellValueTransferBy(String cellValueTransferBy) {
        this.cellValueTransferBy = cellValueTransferBy;
    }

    public String getWidth() {
        return width;
    }

    public void setWidth(String width) {
        this.width = width;
    }

    public String getDataSetId() {
        return dataSetId;
    }

    public void setDataSetId(String dataSetId) {
        this.dataSetId = dataSetId;
    }

    public String getDbColumnName() {
        return dbColumnName;
    }

    public void setDbColumnName(String dbColumnName) {
        this.dbColumnName = dbColumnName;
    }

    public String getTitleComment() {
        return titleComment;
    }

    public void setTitleComment(String titleComment) {
        this.titleComment = titleComment;
    }

    public boolean isVisible() {
        return visible;
    }

    public void setVisible(boolean visible) {
        this.visible = visible;
    }
}
