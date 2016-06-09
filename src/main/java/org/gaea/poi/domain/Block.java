package org.gaea.poi.domain;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 块。对应的是excel中的一个数据区块。一般一个块对应数据库一张数据表。
 * <p/>
 * 合并Block和GaeaPoiResultSet 2016-6-7 11:35:03
 * <p>DataSetResult，是一个特殊的概念。表示一个sheet解析后，其中的一个块。<br/>
 * 这个块包含特定的定义（特别是对应某个表（只能一个）），和相关的一系列的field的定义。<p/>
 * DataSetResult的存在，是为了一个Excel可以导入多个数据表。一份Excel的数据，存在多个表的关联数据。
 * <p/>
 * 例如：导入产品数据，同时还有对应的库存。一行数据就包含了产品信息、库存信息两个表。
 * <p/>
 * 结果集可以嵌套。一般来说，嵌套了的话，就可以无视 sheetDefine, fieldDefines几个字段。</p>
 *
 * Created by iverson on 2016-6-6 16:33:59.
 */
public class Block<T> {
    private String id;
    private String table;// 块对应的数据表
    private String entityClass;// 数据对应的bean全名。例如：com.abc.domain.UserEntity
    private ExcelSheet sheetDefine;
    private List<ExcelField> fieldDefines;
    private Map<String,Field> fieldMap = new LinkedHashMap<String, Field>();// key ： XML定义的name（或Excel定义的name）
    private List<T> data;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getTable() {
        return table;
    }

    public void setTable(String table) {
        this.table = table;
    }

    public String getEntityClass() {
        return entityClass;
    }

    public void setEntityClass(String entityClass) {
        this.entityClass = entityClass;
    }

    public ExcelSheet getSheetDefine() {
        return sheetDefine;
    }

    public void setSheetDefine(ExcelSheet sheetDefine) {
        this.sheetDefine = sheetDefine;
    }

    public List<ExcelField> getFieldDefines() {
        return fieldDefines;
    }

    public void setFieldDefines(List<ExcelField> fieldDefines) {
        this.fieldDefines = fieldDefines;
    }

    public Map<String, Field> getFieldMap() {
        return fieldMap;
    }

    public void setFieldMap(Map<String, Field> fieldMap) {
        this.fieldMap = fieldMap;
    }

    public List<T> getData() {
        return data;
    }

    public void setData(List<T> data) {
        this.data = data;
    }
}
