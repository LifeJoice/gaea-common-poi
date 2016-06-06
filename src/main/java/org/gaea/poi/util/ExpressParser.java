package org.gaea.poi.util;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.lang3.StringUtils;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.config.GaeaPoiDefinition;
import org.gaea.poi.domain.GaeaPoiFieldDefine;
import org.gaea.util.GaeaPropertiesReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Component;

import java.io.IOException;

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
    private ObjectMapper mapper = new ObjectMapper();

    // 可以把整个批注都传进来，本方法只会解释可以识别的
    public GaeaPoiFieldDefine parseField(String excelRemark) throws ValidationFailedException {
        if (StringUtils.isEmpty(excelRemark)
                || (StringUtils.indexOf(excelRemark, cacheProperties.get(GaeaPoiDefinition.FIELD_DEFINE_BEGIN)) < 0)
                || (StringUtils.indexOf(excelRemark, cacheProperties.get(GaeaPoiDefinition.FIELD_DEFINE_END)) < 0)) {
            throw new IllegalArgumentException("GAEA FIELD定义为空或不完整 : "+excelRemark);
        }
        String strExpress = StringUtils.substringBetween(excelRemark,
                cacheProperties.get(GaeaPoiDefinition.FIELD_DEFINE_BEGIN),
                cacheProperties.get(GaeaPoiDefinition.FIELD_DEFINE_END));
        if(StringUtils.isEmpty(strExpress)){
            throw new IllegalArgumentException("GAEA FIELD定义为空 : "+excelRemark);
        }
        GaeaPoiFieldDefine fieldDef = null;
        try {
            fieldDef = mapper.readValue(strExpress,GaeaPoiFieldDefine.class);
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
}