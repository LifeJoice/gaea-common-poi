package org.gaea.poi.service.impl;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.config.GaeaPoiDefinition;
import org.gaea.poi.domain.*;
import org.gaea.poi.domain.Workbook;
import org.gaea.poi.service.ExcelDefineService;
import org.gaea.poi.util.ExpressParser;
import org.gaea.poi.xml.GaeaPoiXmlConfigParser;
import org.gaea.util.BeanUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by iverson on 2017/5/10.
 */
@Service
@Transactional(propagation = Propagation.NEVER)
public class ExcelDefineServiceImpl implements ExcelDefineService {
    private final Logger logger = LoggerFactory.getLogger(ExcelDefineServiceImpl.class);
    @Autowired
    private ExpressParser expressParser;
    @Autowired
    private GaeaPoiXmlConfigParser gaeaPoiXmlConfigParser;

    /**
     * 根据输入的excel input stream，获取对应的Gaea定义（该excel是自描述的）。然后获取系统缓存的定义模板，转换后返回Gaea定义的Workbook。
     *
     * 重构。把定义相关的，放到这个service中。
     * copy from ExcelReader.getWorkbook by Iverson 2017-5-12
     * @param apacheWorkbook    apache的Workbook对象。主要用于获取第一个sheet的第一行，获取模板定义和列定义。
     * @return
     * @throws ValidationFailedException
     */
    public Workbook getWorkbookDefine(org.apache.poi.ss.usermodel.Workbook apacheWorkbook) throws ValidationFailedException {
        Workbook result = null;
//        result.setBlockList(new ArrayList<Block>());
        Map<Integer, Field> columnFieldDefMap = new HashMap<Integer, Field>();
//        try {
//            //根据上述创建的输入流 创建工作簿对象ZA
//            org.apache.poi.ss.usermodel.Workbook wb = WorkbookFactory.create(excelIS);
            if(apacheWorkbook==null || apacheWorkbook.getNumberOfSheets()<1){
                throw new ValidationFailedException("Workbook或者sheet为空，无法获取Gaea定义！");
            }
            //得到第一页 sheet
            //页Sheet是从0开始索引的。默认先支持一个sheet。
            org.apache.poi.ss.usermodel.Sheet sheet = apacheWorkbook.getSheetAt(0);
            Workbook xmlWorkbook = null;
            for (Row row : sheet) {
                // 暂时只在第一行定义导入表达式
                if (row.getRowNum() == GaeaPoiDefinition.GAEA_DEFINE_ROW) {
                    Block block = null;// 所有的数据都关联到块，所以必须先有block定义
                    //遍历row中的所有方格
                    for (Cell cell : row) {
                        // 获取Gaea通用导入的Sheet的定义
                        ExcelSheet excelSheet = getSheetDefine(cell);
                        if (excelSheet != null) {
                            // Workbook的id，就是excelSheet的id，即#GAEA_DEF__BEGIN[]的根id
                            // 把XML配置的workbook信息合并到当前结果中（相同ID的）
                            if (StringUtils.isNotEmpty(excelSheet.getId())) {
                                if(result==null){
                                    result = new Workbook();
                                    result.setBlockList(new ArrayList<Block>());
                                }
                                result.setId(excelSheet.getId());
                                xmlWorkbook = gaeaPoiXmlConfigParser.getWorkbook(excelSheet.getId());
                                org.gaea.util.BeanUtils.copyProperties(xmlWorkbook, result);
                            }
                            // 获取块并设置块。
                            if (CollectionUtils.isNotEmpty(excelSheet.getBlockList())) {
                                // 目前只支持一个sheet
                                block = blockParse(result, excelSheet.getBlockList().get(0));
                                result.getBlockList().add(block);
                            }
                        }
                        // 块不为空，数据才能关联到块。块关联到表！
//                        String myExcelRemark = getCellComment(cell);
                        if (block != null) {
                            ExcelField excelField = getField(cell);
                            // 如果该列不存在定义，跳过
                            if(excelField==null){
                                continue;
                            }
                            if (block.getFieldDefines() == null) {
                                block.setFieldDefines(new ArrayList<ExcelField>());
                            }
                            block.getFieldDefines().add(excelField);
                            // 合并Excel定义和XML定义
                            combineField(block, excelField);
                            // 放入 columnIndex：定义的map中。后面获取值需要根据columnIndex获取使用Field
                            columnFieldDefMap.put(cell.getColumnIndex(), (Field) block.getFieldMap().get(excelField.getName()));
                            // 获取对应的field定义。同样，本来这里也应该是不同block不一样的。现在暂未实现。
//                            ExcelField excelField = expressParser.parseField(myExcelRemark);
//                            excelField.setColumnIndex(cell.getColumnIndex());
//                            if (block.getFieldDefines() == null) {
//                                block.setFieldDefines(new ArrayList<ExcelField>());
//                            }
//                            block.getFieldDefines().add(excelField);
                        }
//                    }
                    }
                }
            }
//        } catch (IOException e) {
//            logger.error("创建Apache Workbook错误。", e);
//        } catch (InvalidFormatException e) {
//            logger.error("创建Apache Workbook错误。", e);
//        }
        return result;
    }
    /**
     * 获取一个cell里，定义的全sheet的内容。
     * @param cell
     * @return
     * @throws ValidationFailedException
     */
    public ExcelSheet getSheetDefine(Cell cell) throws ValidationFailedException {
        String myExcelRemark = null;
        ExcelSheet excelSheet = null;
        if (cell.getCellComment() != null) {
            myExcelRemark = cell.getCellComment().getString().getString();
            excelSheet = expressParser.parseSheet(myExcelRemark);
        }
        return excelSheet;
    }

    public Block blockParse(Workbook workbook, ExcelBlock excelBlock) {
        Block block = null;
        // 合并XML的块配置和Excel文件的块配置信息。
        if (workbook.getBlockList() == null) {
            workbook.setBlockList(new ArrayList<Block>());
        }
        if (workbook.getBlockList() != null) {
            if (workbook.getBlockList().isEmpty()) {
                block = new Block();
//                workbook.getBlockList().add(block);
            } else {
                block = workbook.getBlockList().get(0);
            }
            org.gaea.util.BeanUtils.copyProperties(excelBlock, block);
        }
        return block;
    }

    public String getCellComment(Cell cell) {
        String excelRemark = null;
        if (cell.getCellComment() != null) {
            excelRemark = cell.getCellComment().getString().getString();
        }
        return excelRemark;
    }

    /**
     * 从Gaea定义（sheet）中，提取对应的字段的定义。
     * @param gaeaDefSheet
     * @param fieldDefMap
     * @return
     * @throws ValidationFailedException
     */
    public Map<Integer, Field> getFieldsDefine(Sheet gaeaDefSheet, Map<String, Field> fieldDefMap) throws ValidationFailedException {
        if(gaeaDefSheet!=null){
            Row row = gaeaDefSheet.getRow(GaeaPoiDefinition.GAEA_DEFINE_ROW);
            return getFieldsDefine(row, fieldDefMap);
        }
        return null;
    }
    /**
     * 这个一般是针对Excel第一行操作！
     * <p>
     *     获取行里每一列的定义. 如果还传入了fieldDefMap(很可能来自XML定义等), 则合并两者定义.
     * </p>
     * <p>如果Excel里面没有Gaea定义，则会略过！</p>
     * @param row
     * @param fieldDefMap    可以为空。为空，即以Excel里面的Gaea定义为准！
     * @return
     * @throws ValidationFailedException
     */
    public Map<Integer, Field> getFieldsDefine(Row row, Map<String, Field> fieldDefMap) throws ValidationFailedException {
        Map<Integer, Field> result = new HashMap<Integer, Field>();

        for(Cell cell: row) {
            // 获取单元格备注，提取定义
            ExcelField excelField = getField(cell);
            if(excelField!=null){
                if(StringUtils.isEmpty(excelField.getName())){
                    throw new ValidationFailedException(MessageFormat.format("Excel的Field定义的name不允许为空！row:{0} column:{1}",0,cell.getColumnIndex()).toString());
                }
                // 是否有传入的列定义
                Field field = null;
                if(fieldDefMap!=null) {
                    field = fieldDefMap.get(excelField.getName());
                }
                // 没有则创建
                if(field==null){
                    field = toField(excelField);
                }
                // 刷新field的columnIndex
                field.setColumnIndex(cell.getColumnIndex());
                // 放入结果
                result.put(cell.getColumnIndex(),field);
            }
        }
        return result;
    }

    /**
     * 从Excel的单元格的备注中，提取Gaea框架的定义。
     * @param cell
     * @return  如果没有匹配的语法、定义，返回空。
     * @throws ValidationFailedException
     */
    public ExcelField getField(Cell cell) throws ValidationFailedException {
        String myExcelRemark = getCellComment(cell);
        // 获取对应的field定义。同样，本来这里也应该是不同block不一样的。现在暂未实现。
        ExcelField excelField = expressParser.parseField(myExcelRemark);
        if(excelField==null){
            return null;
        }
        excelField.setColumnIndex(cell.getColumnIndex());
        return excelField;
    }

    public Field toField(ExcelField excelField){
        if(excelField==null){
            return null;
        }
        Field field = new Field();
        BeanUtils.copyProperties(excelField,field);
        return field;
    }

    /**
     * 合并Excel定义和XML定义.
     * <p>
     *     检查block的fieldMap中有没有给定的excelField.如果没有，新增,并把ExcelField的值复制; 如果有,用ExcelField的值覆盖. <br/>
     *     最后刷新block中的fieldMap中同名(name)的field.
     * </p>
     * @param block
     * @param excelField
     */
    public void combineField(Block block, ExcelField excelField) {
        boolean hasXmlFieldDef = false;
        Map<String, Field> fieldMap = block.getFieldMap();
        Field f = fieldMap.get(excelField.getName());// 获取是否有对应的XML定义的field
        // 如果XML定义和Excel定义同时存在，Excel定义相关值覆盖XML定义（但XML是父级，一些关键值EXCEL是覆盖不了的）
        if (f == null) {
            f = new Field();
        }
        org.gaea.util.BeanUtils.copyProperties(excelField, f);
        // 如果没有找到XML定义和Excel定义匹配的，创建一个，并放入block中
        if (!hasXmlFieldDef) {
            block.getFieldMap().put(excelField.getName(), f);
        }
    }
}
