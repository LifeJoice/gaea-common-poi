package org.gaea.poi.cache;

import org.gaea.cache.CacheFactory;
import org.gaea.cache.CacheOperator;
import org.gaea.exception.SysInitException;
import org.gaea.poi.config.GaeaPoiDefinition;
import org.gaea.poi.domain.ExcelTemplate;
import org.gaea.util.GaeaPropertiesReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Component;

/**
 * 这个是Gaea POI包的通用Cache入口。
 * 但它还是依赖CacheFactory。CacheFactory只是一个接口。它需要最终使用方给它注入一个实现。
 * <p>
 * <b>重要！</b><br/>
 * <b>所以，最终引用Gaea POI的系统，必须有CacheFactory的实现，并且托管给Spring！</b>
 * </p>
 * <p/>
 * Created by iverson on 2016/11/5.
 */
@Component
public class GaeaPoiCache {
    private final Logger logger = LoggerFactory.getLogger(GaeaPoiCache.class);
    @Autowired
    private CacheFactory cacheFactory;
    @Autowired
    @Qualifier("gaeaPOIProperties")
    private GaeaPropertiesReader cacheProperties;

    public ExcelTemplate getExcelTemplate(String key) throws SysInitException {
        String redisRootKey = cacheProperties.get(GaeaPoiDefinition.REDIS_EXCEL_EXPORT_TEMPLATE);
        CacheOperator cacheOperator = cacheFactory.getCacheOperator();
        ExcelTemplate excelTemplate = cacheOperator.getHashValue(redisRootKey, key, ExcelTemplate.class);
        if (excelTemplate == null) {
            logger.debug("获取缓存模板失败！redis root key:{} hash key:{}", redisRootKey, key);
        }
        return excelTemplate;
    }
}
