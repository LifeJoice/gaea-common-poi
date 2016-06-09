package org.gaea.poi.xml;

import org.apache.commons.lang3.StringUtils;
import org.gaea.exception.InvalidDataException;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.domain.Block;
import org.gaea.poi.domain.Field;
import org.gaea.poi.domain.Workbook;
import org.gaea.util.GaeaPropertiesReader;
import org.gaea.util.GaeaXmlUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.util.ResourceUtils;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.annotation.PostConstruct;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 负责通用excel导入的XML配置信息的处理的核心。<p/>
 * 这个没有用注解让Spring自动扫描，因为初始化的时候会初始化数据集。因此构建bean的时候，需要XML配置文件路径。<p/>
 * 使用：<br/>
 * 通过XML配置bean。
 * Created by iverson on 2016-6-6 16:31:50.
 */
public class GaeaPoiXmlConfigParser {

    private final Logger logger = LoggerFactory.getLogger(GaeaPoiXmlConfigParser.class);

    public GaeaPoiXmlConfigParser(String filePath) {
        this.filePath = filePath;
        this.workbookMap = new ConcurrentHashMap<String, Workbook>();
    }

    private String filePath;
    // 系统初始化后，所有配置的DataSet都会加载进来，并不再加载
    private final ConcurrentHashMap<String, Workbook> workbookMap;
//    @Autowired(required = false)
//    @Qualifier("cachePropReader")
//    private GaeaPropertiesReader cacheProperties;
//    @Autowired(required = false)
//    private GaeaCacheProcessor gaeaCacheProcessor;

//    public GaeaDataSetResolver() {
//        this.workbookMap = new ConcurrentHashMap<String, GaeaDataSet>();
//    }

    public Workbook getWorkbook(String id) {
        return workbookMap.get(id);
    }

    /**
     * 根据配置XML文件，读取XML中的导入相关配置（workbook等）
     *
     * @throws ValidationFailedException
     */
    @PostConstruct
    public void init() throws ValidationFailedException {
        if (StringUtils.isEmpty(filePath)) {
            logger.warn("filePath为空！无法进行Gaea Poi的初始化操作。");
        }
        /* 读取XML文件，把DataSet读取和转换处理。 */
        readAndParseXml();
        /* 完成DataSet的XML的加载，接下来缓存 */
//        cacheDataSets();
    }

    private void readAndParseXml() throws ValidationFailedException {
        DocumentBuilder db = null;
        Node document = null;
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
//        Resource resource = springApplicationContext.getResource(viewSchemaPath);
        try {
            File dsXmlFile = ResourceUtils.getFile(filePath);
            db = dbf.newDocumentBuilder();
            // document是整个XML schema
            document = db.parse(dsXmlFile);
            // 寻找根节点<poi-definition>
            Node rootNode = getRootNode(document);

            NodeList nodes = rootNode.getChildNodes();
            //遍历所有子节点<workbook>
            for (int i = 0; i < nodes.getLength(); i++) {
                Node workbookNode = nodes.item(i);
                // xml解析会把各种换行符等解析成元素。统统跳过。
                if (!(workbookNode instanceof Element)) {
                    continue;
                }
                if (GaeaPoiXmlSchemaDefinition.POI_WORKBOOK_NODE_NAME.equalsIgnoreCase(workbookNode.getNodeName())) {
                    Workbook workbook = parseWorkbook(workbookNode);
                    if (workbook == null || StringUtils.isEmpty(workbook.getId())) {
                        logger.warn("格式不正确。对应的DataSet为空或缺失id！" + workbookNode.toString());
                        continue;
                    }
                    workbookMap.put(workbook.getId(), workbook);
                } else {
                    logger.warn("Gaea Poi配置Xml schema中包含错误定义。包含不可识别信息: <" + workbookNode.getNodeName() + ">");
                }
            }
            /* 完成DataSet的XML的加载，接下来缓存 */
//            cacheDataSets();
        } catch (FileNotFoundException e) {
            logger.error("加载GaeaPOI的XML配置文件错误。File path:" + filePath, e);
        } catch (ParserConfigurationException e) {
            logger.error("解析GaeaPOI的XML配置文件错误。File path:" + filePath, e);
        } catch (IOException e) {
            logger.error("解析GaeaPOI的XML配置文件发生IO错误。File path:" + filePath, e);
        } catch (SAXException e) {
            logger.error("解析GaeaPOI的XML配置文件错误。File path:" + filePath, e);
        } catch (InvalidDataException e) {
            logger.error(e.getMessage(), e);
        }
    }

    /**
     * 把从XML文件中读出来的DataSet缓存起来。
     */
//    private void cacheDataSets() {
//        if (workbookMap != null && workbookMap.size() > 0) {
//            String rootKey = cacheProperties.get(GaeaDataSetDefinition.GAEA_DATASET_SCHEMA);
////            for(GaeaDataSet ds:workbookMap.values()){
////                dsMap.put(ds.getId(),ds.getSql());
////            }
//            gaeaCacheProcessor.put(rootKey, workbookMap);
//        }
//    }

    /**
     * 转换XML文件中的单个DataSet
     *
     * @param workbookNode
     * @return
     * @throws InvalidDataException
     */
    private Workbook parseWorkbook(Node workbookNode) throws InvalidDataException, ValidationFailedException {
        Workbook workbook = new Workbook();
        Element workbookElement = (Element) workbookNode;
        NodeList nodes = workbookElement.getChildNodes();
        // 先自动填充<dataset>的属性
        try {
            workbook = GaeaXmlUtils.copyAttributesToBean(workbookNode, workbook, Workbook.class);
        } catch (Exception e) {
            String errorMsg = "自动转换XML元素<workbook>的属性错误！";
            throw new InvalidDataException(errorMsg, e);
        }
        for (int i = 0; i < nodes.getLength(); i++) {
            Node blockNode = nodes.item(i);
            // xml解析会把各种换行符等解析成元素。统统跳过。
            if (!(blockNode instanceof Element)) {
                continue;
            }
            if (GaeaPoiXmlSchemaDefinition.POI_WORKBOOK_BLOCK_NODE_NAME.equalsIgnoreCase(blockNode.getNodeName())) {
                Block block = parseBlock(blockNode);
//                Block block = new Block();
//                try {
//                    block = GaeaXmlUtils.copyAttributesToBean(blockNode, block, Block.class);
//                } catch (Exception e) {
//                    String errorMsg = "自动转换XML元素<block>的属性错误！";
//                    throw new InvalidDataException(errorMsg, e);
//                }
                if (block != null) {
                    if (workbook.getBlockList() == null) {
                        workbook.setBlockList(new ArrayList<Block>());
                    }
                    workbook.getBlockList().add(block);
                }
            } else {
                logger.warn("Gaea Poi配置Xml schema中包含错误定义。包含不可识别信息: <" + blockNode.getNodeName() + ">");
            }
        }
        return workbook;
    }

    /**
     * 转换XML <block>的属性和所有的子<field>至bean对象.
     * @param blockNode
     * @return
     * @throws InvalidDataException
     */
    private Block parseBlock(Node blockNode) throws InvalidDataException, ValidationFailedException {
        Block block = new Block();
        Element blockElement = (Element) blockNode;
        NodeList fieldNodes = blockElement.getChildNodes();
        try {
            block = GaeaXmlUtils.copyAttributesToBean(blockNode, block, Block.class);
        } catch (Exception e) {
            String errorMsg = "自动转换XML元素<block>的属性错误！";
            throw new InvalidDataException(errorMsg, e);
        }
        for (int i = 0; i < fieldNodes.getLength(); i++) {
            Node fieldNode = fieldNodes.item(i);
            // xml解析会把各种换行符等解析成元素。统统跳过。
            if (!(fieldNode instanceof Element)) {
                continue;
            }
            if (GaeaPoiXmlSchemaDefinition.POI_WORKBOOK_FIELD_NODE_NAME.equalsIgnoreCase(fieldNode.getNodeName())) {
                Field field = parseField(fieldNode);
                if(StringUtils.isEmpty(field.getName())){
                    throw new ValidationFailedException("XML定义的<field>的name属性不允许为空！");
                }
                if (field != null) {
//                    if (block.getFieldList() == null) {
//                        block.setFieldList(new ArrayList<Field>());
//                    }
                    block.getFieldMap().put(field.getName(),field);
                }
            } else {
                logger.warn("Gaea Poi配置Xml schema中包含错误定义。包含不可识别信息: <" + fieldNode.getNodeName() + ">");
            }
        }
        return block;
    }

    /**
     * 转换XML <field>的属性至bean对象.
     * @param fieldNode
     * @return
     * @throws InvalidDataException
     */
    private Field parseField(Node fieldNode) throws InvalidDataException {
        Field field = new Field();
        try {
            field = GaeaXmlUtils.copyAttributesToBean(fieldNode, field, Field.class);
        } catch (Exception e) {
            String errorMsg = "自动转换XML元素<field>的属性错误！";
            throw new InvalidDataException(errorMsg, e);
        }
        return field;
    }

    /**
     * 获取整个XML文档的根节点。
     *
     * @param document
     * @return
     * @throws ValidationFailedException
     */
    private Node getRootNode(Node document) throws ValidationFailedException {
        NodeList nodes = document.getChildNodes();
        Node rootNode = null; // 这个应该是<poi-definition>
        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            // xml解析会把各种换行符等解析成元素。统统跳过。
            if (!(node instanceof Element)) {
                continue;
            }
            if (GaeaPoiXmlSchemaDefinition.POI_ROOT_NODE.equalsIgnoreCase(node.getNodeName())) {
                rootNode = node;
                break;
            }
        }
        if (rootNode == null) {
            logger.warn("Gaea Poi 导入XML Schema根节点为空。加载失败。");
        }
        return rootNode;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }
}
