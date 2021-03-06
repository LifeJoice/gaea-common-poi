package org.gaea.poi.config;

/**
 * Created by iverson on 2016-6-4 16:10:12.
 */
public interface GaeaPoiDefinition {
    public static final String POI_DEFAULT_DATE_FORMAT = "gaea.poi.default.dateformat";
    public static final String POI_DEFAULT_TIME_FORMAT = "gaea.poi.default.timeformat";
    public static final String POI_DEFAULT_DATETIME_FORMAT = "gaea.poi.default.datetimeformat";
    public static final String DEFINE_BEGIN = "gaea.poi.template.define.begin";
    public static final String DEFINE_END = "gaea.poi.template.define.end";
    public static final String FIELD_DEFINE_BEGIN = "gaea.poi.template.field.define.begin";
    public static final String FIELD_DEFINE_END = "gaea.poi.template.field.define.end";
    /**
     * =============================================================================================================================================
     *                                                              REDIS
     * =============================================================================================================================================
     */
    public static final String REDIS_EXCEL_EXPORT_TEMPLATE = "gaea.redis.excel.export.template"; // excel export template的prop key
    /**
     * =============================================================================================================================================
     *                                                              普通定义
     * =============================================================================================================================================
     */
    public static final int GAEA_DEFINE_SHEET = 1; // 默认第二个是Gaea的excel模板定义sheet
    public static final int GAEA_DEFINE_ROW = 0; // 默认第一行是Gaea的excel模板定义
    public static final int EXCEL_TITLE_ROW = 0; // 第二行就是一般的Excel表普通title(表示会略过, 不会当数据读取)
}
