package org.gaea.poi.domain;

import java.util.List;

/**
 * DataSetResult，是一个特殊的概念。表示一个sheet解析后，其中的一个块。<br/>
 * 这个块包含特定的定义（特别是对应某个表（只能一个）），和相关的一系列的field的定义。<p/>
 * DataSetResult的存在，是为了一个Excel可以导入多个数据表。一份Excel的数据，存在多个表的关联数据。
 * <p/>
 * 例如：导入产品数据，同时还有对应的库存。一行数据就包含了产品信息、库存信息两个表。
 * <p/>
 * 结果集可以嵌套。一般来说，嵌套了的话，就可以无视 sheetDefine, fieldDefines几个字段。
 * Created by iverson on 2016-6-4 16:48:09.
 */
public class GaeaPoiResultSet<T> {
    private GaeaPoiSheetDefine sheetDefine;
    private List<GaeaPoiFieldDefine> fieldDefines;
    private List<T> data;

    public GaeaPoiSheetDefine getSheetDefine() {
        return sheetDefine;
    }

    public void setSheetDefine(GaeaPoiSheetDefine sheetDefine) {
        this.sheetDefine = sheetDefine;
    }

    public List<GaeaPoiFieldDefine> getFieldDefines() {
        return fieldDefines;
    }

    public void setFieldDefines(List<GaeaPoiFieldDefine> fieldDefines) {
        this.fieldDefines = fieldDefines;
    }

    public List<T> getData() {
        return data;
    }

    public void setData(List<T> data) {
        this.data = data;
    }
}
