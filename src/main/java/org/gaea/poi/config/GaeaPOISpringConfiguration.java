package org.gaea.poi.config;

import org.gaea.util.GaeaPropertiesReader;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.stereotype.Component;

import java.util.Arrays;

/**
 * Created by iverson on 2016/6/4.
 */
@Component
@Configuration
public class GaeaPoiSpringConfiguration {
    public static final String DEFAULT_PROPERTIES = "gaea-poi-config.properties";
    @Bean(name = "gaeaPOIProperties")
    public GaeaPropertiesReader promotionEventPublisherServiceExporter() {
        GaeaPropertiesReader cacheProperties = new GaeaPropertiesReader(Arrays.asList("classpath://"+DEFAULT_PROPERTIES));
        return cacheProperties;
    }
}
