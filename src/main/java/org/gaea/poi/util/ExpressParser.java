package org.gaea.poi.util;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.lang3.StringUtils;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.config.GaeaPoiDefinition;
import org.gaea.poi.domain.ExcelBlock;
import org.gaea.poi.domain.ExcelField;
import org.gaea.poi.domain.ExcelSheet;
import org.gaea.poi.domain.Field;
import org.gaea.util.GaeaPropertiesReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Gaea POI针对EXCEL里面表达式的通用解析器。
 * Created by iverson on 2016-6-4 14:58:21.
 */
@Component
public class ExpressParser {
    private final Logger logger = LoggerFactory.getLogger(ExpressParser.class);
    @Autowired
    @Qualifier("gaeaPOIProperties")
    private GaeaPropertiesReader cacheProperties;
    private static ObjectMapper mapper = new ObjectMapper();

    // 把一个单元格上的批注，抽取出Sheet（包含Block）的定义
    public ExcelSheet parseSheet(String excelRemark) throws ValidationFailedException {
        if (StringUtils.isEmpty(excelRemark)
                || (StringUtils.indexOf(excelRemark, cacheProperties.get(GaeaPoiDefinition.DEFINE_BEGIN)) < 0)
                || (StringUtils.indexOf(excelRemark, cacheProperties.get(GaeaPoiDefinition.DEFINE_END)) < 0)) {
            return null;
        }
        /* 提取表达式部分 */
        String strExpress = StringUtils.substringBetween(excelRemark,
                cacheProperties.get(GaeaPoiDefinition.DEFINE_BEGIN),
                cacheProperties.get(GaeaPoiDefinition.DEFINE_END));
        if(StringUtils.isEmpty(strExpress)){
            throw new IllegalArgumentException("GAEA FIELD定义为空 : "+excelRemark);
        }
        ExcelSheet excelSheet = null;
        try {
            excelSheet = mapper.readValue(strExpress,ExcelSheet.class);
        } catch (IOException e) {
            e.printStackTrace();
            throw new ValidationFailedException("POI sheet表达式( #GAEA_DEF__BEGIN[...] )解析错误 : "+strExpress);
        }
        return excelSheet;
    }
    // 可以把整个批注都传进来，本方法只会解释可以识别的
    public ExcelField parseField(String excelRemark) throws ValidationFailedException {
        if (StringUtils.isEmpty(excelRemark)
                || (StringUtils.indexOf(excelRemark, cacheProperties.get(GaeaPoiDefinition.FIELD_DEFINE_BEGIN)) < 0)
                || (StringUtils.indexOf(excelRemark, cacheProperties.get(GaeaPoiDefinition.FIELD_DEFINE_END)) < 0)) {
//            throw new IllegalArgumentException("GAEA FIELD定义为空或不完整 : "+excelRemark);
            return null;
        }
        String strExpress = StringUtils.substringBetween(excelRemark,
                cacheProperties.get(GaeaPoiDefinition.FIELD_DEFINE_BEGIN),
                cacheProperties.get(GaeaPoiDefinition.FIELD_DEFINE_END));
        if(StringUtils.isEmpty(strExpress)){
            throw new IllegalArgumentException("GAEA FIELD定义为空 : "+excelRemark);
        }
        ExcelField fieldDef = null;
        try {
            fieldDef = mapper.readValue(strExpress,ExcelField.class);
        } catch (IOException e) {
            throw new ValidationFailedException("POI field表达式解析错误 : "+strExpress);
        }
        return fieldDef;
    }

    public static void main(String[] args) {
        String excelRemark = "作者:\n" +
                "#GAEA_DEF_BEGIN[\n" +
                "{\n" +
                "  \"tableAlias\" : \"userAttendance\"\n" +
                "}\n" +
                "]GAEA_DEF_END#\n" +
                "\n" +
                "#GAEA_DEF_FIELD[\n" +
                "{\n" +
                "  \"name\" : \"username\"\n" +
                "}\n" +
                "]GAEA_DEF_FIELD#";
//        String searchStr = "#GAEA_DEF_FIELD[";
        String searchStr = "]GAEA_DEF_FIELD#";
        System.out.println(StringUtils.indexOf(excelRemark,searchStr));
    }

    /**
     * 根据输入项，创建一个ExcelSheet对象。
     * @param excelTemplateId
     * @return
     * @throws ValidationFailedException
     */
    public static ExcelSheet createSheet(String excelTemplateId) throws ValidationFailedException {
        if(StringUtils.isEmpty(excelTemplateId)){
            throw new ValidationFailedException("模板id为空，无法创建Gaea ExcelSheet定义对象！");
        }
        ExcelSheet excelSheet = new ExcelSheet();
        excelSheet.setId(excelTemplateId);
        ExcelBlock excelBlock = new ExcelBlock();
        // 当前还没有启用block功能，默认给个AUTO
        excelBlock.setId("AUTO");
        excelSheet.getBlockList().add(excelBlock);
        return excelSheet;
    }

    /**
     * 创建gaea ExcelSheet定义的模板字符串。
     * 把excelSheet定义对象，转换为Excel中可以直接用的Gaea模板定义（字符串）。
     * @param excelSheet
     * @return
     * @throws JsonProcessingException
     */
    public static String createSheetDefine(ExcelSheet excelSheet) throws JsonProcessingException {
        if(excelSheet==null){
            return "";
        }
        StringBuilder result = new StringBuilder();
        result.append(GaeaPoiProperties.get(GaeaPoiDefinition.DEFINE_BEGIN));
        result.append("\n");
        result.append(mapper.writeValueAsString(excelSheet));
        result.append("\n");
        result.append(GaeaPoiProperties.get(GaeaPoiDefinition.DEFINE_END));
        return result.toString();
    }

    /**
     * 把field定义对象，转换为Excel中可以直接用的Gaea模板定义（字符串）。
     *
     * @param field
     * @return
     * @throws JsonProcessingException
     */
    public static String createFieldDefine(Field field) throws JsonProcessingException {
        if(field==null){
            return "";
        }
        // 先转换成gaea的ExcelField
        ExcelField excelField = new ExcelField();
        excelField.setColumnIndex(field.getColumnIndex());
        excelField.setName(field.getName());

        StringBuilder result = new StringBuilder("\n");
        result.append(GaeaPoiProperties.get(GaeaPoiDefinition.FIELD_DEFINE_BEGIN));
        result.append("\n");
        result.append(mapper.writeValueAsString(excelField));
        result.append("\n");
        result.append(GaeaPoiProperties.get(GaeaPoiDefinition.FIELD_DEFINE_END));
        return result.toString();
    }
}
