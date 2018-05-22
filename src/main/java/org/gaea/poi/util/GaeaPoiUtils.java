package org.gaea.poi.util;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.config.GaeaPoiDefinition;
import org.gaea.poi.domain.Field;
import org.gaea.util.GaeaDateTimeUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.LinkedCaseInsensitiveMap;

import java.sql.Timestamp;
import java.text.ParseException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by iverson on 2017/5/16.
 */
public class GaeaPoiUtils {
    private static final Logger logger = LoggerFactory.getLogger(GaeaPoiUtils.class);

    /**
     * 获取excel里一个cell的值。根据类型判断，但最后统一转换为String。
     * 因为POI的接口没有统一返回cell中的值的。
     * <p/>
     * copy from ExcelReaderImpl.getCellStringValue
     *
     * @param cell
     * @param dataType XML定义的读取类型。为空则按Excel单元格类型转换。
     * @return
     */
    public static String getCellStringValue(Cell cell, String dataType) {
        String value = "";
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            value = cell.getStringCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) { // cellType=0, 不只是数字，常规也是这个
            // 如果是日期类型
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                value = cell.getDateCellValue().toString();
                // 如果XML配置了对应的日期类型，则按XML配置转换；否则按系统配置的默认年月日时分秒转换。
                if (Field.DATA_TYPE_DATE.equalsIgnoreCase(dataType)) {
                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATE_FORMAT));
                } else if (Field.DATA_TYPE_TIME.equalsIgnoreCase(dataType)) {
                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_TIME_FORMAT));
                } else if (Field.DATA_TYPE_DATETIME.equalsIgnoreCase(dataType)) {
                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT));
                } else {
                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT));
                }
                return value;
            }
//            value = String.valueOf(cell.getNumericCellValue());
            // 这里要用原始值。否则对于一些值例如：“1”，在Excel显示是“1”，getNumericCellValue读进来就会变成“1.0”
            value = ((XSSFCell) cell).getRawValue();
        }
        return value;
    }

    /**
     * 根据Excel Template定义，如果字段有定义其他类型的，例如数字类，则作转换。再把值回写cell。并且把cell的类型设置为对应的类型。
     * 这样，比如对于一些数值类的，在excel里面才可以直接用公式计算。不用再做转换。
     * <p>
     * 日期类的，兼容输入是长整型的，或者字符型的（格式yyyy-MM-dd）
     * </p>
     * <p/>
     * copy from ExcelExportImpl.parseCellData
     *
     * @param cell
     * @param inValue
     * @param fieldDef
     * @throws ValidationFailedException
     */
    public static void setCellValue(Cell cell, Object inValue, Field fieldDef) throws ValidationFailedException {
        // 默认单元格类型
        cell.setCellType(XSSFCell.CELL_TYPE_STRING);
        String value = String.valueOf(inValue);
        try {
            if (StringUtils.isEmpty(fieldDef.getDataType()) || Field.DATA_TYPE_STRING.equalsIgnoreCase(fieldDef.getDataType())) {
                // default. do nothing.
            } else if (Field.DATA_TYPE_NUMBER.equalsIgnoreCase(fieldDef.getDataType())) {
                cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//                cell.setCellValue(value);
            } else if (Field.DATA_TYPE_DATE.equalsIgnoreCase(fieldDef.getDataType())) {
                String dtPattern = StringUtils.isNotEmpty(fieldDef.getDatetimeFormat()) ? fieldDef.getDatetimeFormat() : GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATE_FORMAT);
                /**
                 * 如果传入的值是整型，先转换成Date，再格式化
                 */
                value = getCellDateTimeValue(inValue, fieldDef, dtPattern);
//                if (NumberUtils.isNumber(value)) {
//                    value = DateFormatUtils.format(new Date(Long.parseLong(value)), GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATE_FORMAT));
//                }
//                // 按单元格定义，转换格式
//                if (StringUtils.isNotEmpty(fieldDef.getDatetimeFormat())) {
//                    // 先按标准日期格式转成Date，再按特定要求转换格式
//                    value = DateFormatUtils.format(DateUtils.parseDate(value, GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATE_FORMAT)), fieldDef.getDatetimeFormat());
//                }
            } else if (Field.DATA_TYPE_TIME.equalsIgnoreCase(fieldDef.getDataType())) {
                value = getCellDateTimeValue(inValue, fieldDef, GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_TIME_FORMAT));
                /**
                 * 如果传入的值是整型，先转换成Date，再格式化
                 */
//                if (NumberUtils.isNumber(value)) {
//                    value = DateFormatUtils.format(new Date(Long.parseLong(value)), GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_TIME_FORMAT));
//                }
//                // 按单元格定义，转换格式
//                if (StringUtils.isNotEmpty(fieldDef.getDatetimeFormat())) {
//                    // 先按标准日期格式转成Date，再按特定要求转换格式
//                    value = DateFormatUtils.format(DateUtils.parseDate(value, GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_TIME_FORMAT)), fieldDef.getDatetimeFormat());
//                }
            } else if (Field.DATA_TYPE_DATETIME.equalsIgnoreCase(fieldDef.getDataType())) {
                value = getCellDateTimeValue(inValue, fieldDef, GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT));
                /**
                 * 如果传入的值是整型，先转换成Date，再格式化
                 */
//                if (NumberUtils.isNumber(value)) {
//                    value = DateFormatUtils.format(new Date(Long.parseLong(value)), GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT));
//                }
//                // 按单元格定义，转换格式
//                if (StringUtils.isNotEmpty(fieldDef.getDatetimeFormat())) {
//                    // 先按标准日期格式转成Date，再按特定要求转换格式
//                    value = DateFormatUtils.format(DateUtils.parseDate(value, GaeaPoiProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT)), fieldDef.getDatetimeFormat());
//                }
            }
            cell.setCellValue(value);
        } catch (NumberFormatException e) {
            throw new ValidationFailedException("单元格数据类型错误，转换失败！value: " + value + " dataType: " + fieldDef.getDataType(), e);
        } catch (ParseException e) {
            throw new ValidationFailedException("单元格数据类型错误，转换日期失败！value: " + value + " dataType: " + fieldDef.getDataType(), e);
        }
    }

    public static String getCellDateTimeValue(Object inValue, Field fieldDef, String dateTimePattern) throws ParseException {
        if (inValue instanceof String && StringUtils.isEmpty((CharSequence) inValue)) {
            return "";
        }
        String result = "";
        /* 如果传入的值是整型，先转换成Date，再格式化 */
        if (inValue instanceof Date || inValue instanceof Timestamp || inValue instanceof java.sql.Date) {
            // 上面几个类型都是java.util.Date的子类，统统强制转Date
            result = DateFormatUtils.format((Date) inValue, dateTimePattern);
        }
        /* 如果传入的值是整型，先转换成Date，再格式化 */
        else if (NumberUtils.isNumber(String.valueOf(inValue))) {
            result = DateFormatUtils.format(new Date(Long.parseLong(String.valueOf(inValue))), dateTimePattern);
        }
        // 按单元格定义，转换格式
        else if (StringUtils.isNotEmpty(fieldDef.getDatetimeFormat())) {
            // 先按标准日期格式转成Date，再按特定要求转换格式
            result = DateFormatUtils.format(DateUtils.parseDate(String.valueOf(inValue), GaeaDateTimeUtils.getDefaultConvertPatterns()), fieldDef.getDatetimeFormat());
        }
        return result;
    }

    /**
     * 由于通过Block获取的FieldMap，是用了db-column-name作为key。这里转换为用name作为map的key
     * @param fieldMap
     * @return
     */
    public static LinkedCaseInsensitiveMap<Field> getNameColumnMap(Map<String, Field> fieldMap) {
        if (MapUtils.isEmpty(fieldMap)) {
            return null;
        }
        LinkedCaseInsensitiveMap<Field> columnMap = new LinkedCaseInsensitiveMap<Field>();
        for (String key : fieldMap.keySet()) {
            Field excelField = fieldMap.get(key);
            if (StringUtils.isEmpty(excelField.getDbColumnName())) {
                logger.debug("XML配置的excel field确实db-column-name，无法转换成DbNameColumnMap.Field name:{}", excelField.getName());
                continue;
            }
            columnMap.put(excelField.getName(), excelField);
        }
        return columnMap;
    }

    /**
     * 获取两个集合相交的Field。其中一个集合是完整的。另一个List，是只有key的list。
     *
     * @param linkFieldsMap          这是一个有序的map。而返回的map的顺序，也保持和这个一样。
     * @param fieldsKeyList
     * @return
     * @throws ValidationFailedException
     */
    public static LinkedCaseInsensitiveMap<Field> getJoinFields(Map<String, Field> linkFieldsMap, List<String> fieldsKeyList) throws ValidationFailedException {
        if (CollectionUtils.isEmpty(fieldsKeyList) || MapUtils.isEmpty(linkFieldsMap)) {
            return null;
        }
        LinkedCaseInsensitiveMap<Field> resultMap = null;
        resultMap = new LinkedCaseInsensitiveMap<Field>();
        // 遍历有序的map，这样返回才能根据有序map的顺序返回
        for (String linkKey : linkFieldsMap.keySet()) {
            for (String key2 : fieldsKeyList) {
                if (linkKey.equalsIgnoreCase(key2)) {
                    resultMap.put(key2, linkFieldsMap.get(key2));
                    break;
                }
            }
        }
        return resultMap;
    }
}
