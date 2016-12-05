package org.gaea.poi.export;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.gaea.data.cache.CacheFactory;
import org.gaea.data.cache.CacheOperator;
import org.gaea.exception.InvalidDataException;
import org.gaea.exception.SysInitException;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.XmlSchemaDefinition;
import org.gaea.poi.config.GaeaPoiDefinition;
import org.gaea.poi.domain.*;
import org.gaea.util.GaeaPropertiesReader;
import org.gaea.util.GaeaXmlUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.core.io.support.ResourcePatternUtils;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.annotation.PostConstruct;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Gaea框架的Excel导出处理器。
 * 负责初始化Gaea POI框架。加载各个模板文件。
 * 这个需要用Spring手动注册，不建议自动托管。因为需要一些启动参数。
 * Created by iverson on 2016-11-1 10:56:28.
 */
public class GaeaExcelExportProcessor {
    private final Logger logger = LoggerFactory.getLogger(GaeaExcelExportProcessor.class);

    @Autowired
    private ResourceLoader resourceLoader;
    @Autowired
    @Qualifier("gaeaPOIProperties")
    private GaeaPropertiesReader cacheProperties;
    @Autowired
    private CacheFactory cacheFactory;
    /**
     * 这个适合读取classpath等路径的文件。例如：classpath:xxx.properties
     * 不适合读取WEB-INF的文件，因为获取不到容器的上下文，无法定位(ServletContext)
     */
    private List<String> resourceLocations = null;

    /**
     * 不建议用这个构造器。基本上用这个是无效的。
     */
    public GaeaExcelExportProcessor() {
    }

    public GaeaExcelExportProcessor(List<String> resourceLocations) {
        this.resourceLocations = resourceLocations;
    }

    /**
     * 在bean初始化后，读取resourceLocations资源。
     */
    @PostConstruct
    public void init() {
        try {
            Resource[] resources = null;
            for (String rsLocation : resourceLocations) {
                Resource[] arrayR = ResourcePatternUtils.getResourcePatternResolver(resourceLoader).getResources(rsLocation);
                resources = ArrayUtils.addAll(resources, arrayR);
            }
            parseAndCache(resources);
        } catch (IOException e) {
            logger.error("初始化GaeaExcelExportProcessor失败。加载Gaea框架通用Excel导出所需相关资源失败。" + e.getMessage(), e);
        } catch (InvalidDataException e) {
            logger.error(e.getMessage(), e);
        } catch (SysInitException e) {
            logger.error(e.getMessage(), e);
        } catch (ValidationFailedException e) {
            logger.error(e.getMessage(), e);
        }
    }

    private void parseAndCache(Resource[] resources) throws InvalidDataException, ValidationFailedException, SysInitException {
        if (ArrayUtils.isNotEmpty(resources)) {
            String redisRootKey = cacheProperties.get(GaeaPoiDefinition.REDIS_EXCEL_EXPORT_TEMPLATE);
            CacheOperator cacheOperator = cacheFactory.getCacheOperator();
            for (Resource r : resources) {
                List<ExcelTemplate> excelTemplateList = parse(r);
                if (CollectionUtils.isNotEmpty(excelTemplateList)) {
                    for (ExcelTemplate excelTemplate : excelTemplateList) {
                        /**
                         * Redis缓存（参考）:
                         * GAEA:EXCEL:EXPORT:TEMPLATE = HashMap< ExcepTemplate id : ExcelTemplate obj >
                         */
                        cacheOperator.putHashValue(redisRootKey, excelTemplate.getId(), excelTemplate);
                    }
                }
            }
        } else {
            logger.debug("GaeaExcelExportProcessor未加载任何excel模板。模板列表为空。");
        }
    }

    /**
     * 处理某一个Excel 模板文件。转换成java对象。
     * @param resource
     * @return
     * @throws ValidationFailedException
     * @throws InvalidDataException
     */
    private List<ExcelTemplate> parse(Resource resource) throws ValidationFailedException, InvalidDataException {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        DocumentBuilder db = null;
        Node document = null;
        List<ExcelTemplate> excelTemplateList = new ArrayList<ExcelTemplate>();
        try {
            db = dbf.newDocumentBuilder();
            // document是整个XML schema
            document = db.parse(resource.getInputStream());
            // 寻找根节点<excel-template>
            NodeList nodes = document.getChildNodes();
            for (int i = 0; i < nodes.getLength(); i++) {
                Node node = nodes.item(i);
                // xml解析会把各种换行符等解析成元素。统统跳过。
                if (!(node instanceof Element)) {
                    continue;
                }
                if (XmlSchemaDefinition.ROOT_NODE.equals(node.getNodeName())) {
                    ExcelTemplate excelTemplate = parseExcelTemplate(node);
                    excelTemplateList.add(excelTemplate);
                }
            }
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return excelTemplateList;
    }

    private ExcelTemplate parseExcelTemplate(Node inNode) throws ValidationFailedException, InvalidDataException {
        ExcelTemplate excelTemplate = new ExcelTemplate();
        excelTemplate = GaeaXmlUtils.copyAttributesToBean(inNode, excelTemplate, ExcelTemplate.class);
        Sheet sheet = new Sheet();
        NodeList nodes = inNode.getChildNodes();
        for (int i = 0; i < nodes.getLength(); i++) {
            Node n = nodes.item(i);
            // xml解析会把各种换行符等解析成元素。统统跳过。
            if (!(n instanceof Element)) {
                continue;
            }
            if (XmlSchemaDefinition.BLOCK_NAME.equals(n.getNodeName())) {
                // 暂时不支持多个sheet。所以默认初始化一个。
                if (excelTemplate.getExcelSheetList() == null) {
                    List<Block> blockList = new ArrayList<Block>();
                    sheet.setBlockList(blockList);
                    List<Sheet> list = new ArrayList<Sheet>();
                    list.add(sheet);
                    excelTemplate.setExcelSheetList(list);
                }
                Block block = parseBlock(n);
                sheet.getBlockList().add(block);
            }
        }
        return excelTemplate;
    }

    private Block parseBlock(Node inNode) throws InvalidDataException {
        Block block = new Block();
        block = GaeaXmlUtils.copyAttributesToBean(inNode, block, Block.class);
        NodeList nodes = inNode.getChildNodes();
        for (int i = 0; i < nodes.getLength(); i++) {
            Node n = nodes.item(i);
            // xml解析会把各种换行符等解析成元素。统统跳过。
            if (!(n instanceof Element)) {
                continue;
            }
            if (XmlSchemaDefinition.FIELD_NAME.equals(n.getNodeName())) {
                Field field = parseField(n);
                block.getFieldMap().put(field.getDbColumnName(), field);
            }
        }
        return block;
    }

    private Field parseField(Node inNode) throws InvalidDataException {
        Field field = new Field();
        field = GaeaXmlUtils.copyAttributesToBean(inNode, field, Field.class);
        return field;
    }
}
