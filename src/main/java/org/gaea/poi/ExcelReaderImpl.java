package org.gaea.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.config.GaeaPoiDefinition;
import org.gaea.poi.domain.*;
import org.gaea.poi.service.ExcelDefineService;
import org.gaea.poi.util.ExpressParser;
import org.gaea.poi.util.GaeaPoiUtils;
import org.gaea.poi.xml.GaeaPoiXmlConfigParser;
import org.gaea.util.GaeaPropertiesReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by iverson on 2016-6-3 17:58:00
 */
@Component
public class ExcelReaderImpl implements ExcelReader {
    private final Logger logger = LoggerFactory.getLogger(ExcelReaderImpl.class);
    @Autowired
    private ExpressParser expressParser;
    @Autowired
    private GaeaPoiXmlConfigParser gaeaPoiXmlConfigParser;
    @Autowired
    @Qualifier("gaeaPOIProperties")
    private GaeaPropertiesReader cacheProperties;
    @Autowired
    private ExcelDefineService excelDefineService;

    public org.gaea.poi.domain.Workbook getWorkbook(InputStream fileIS) throws ValidationFailedException {
        org.gaea.poi.domain.Workbook result = null;
//        result.setBlockList(new ArrayList<Block>());
        Map<Integer, Field> columnFieldDefMap = new HashMap<Integer, Field>();
        try {
            //根据上述创建的输入流 创建工作簿对象ZA
            Workbook wb = WorkbookFactory.create(fileIS);
            //得到第一页 sheet
            //页Sheet是从0开始索引的。默认先支持一个sheet。
            Sheet sheet = wb.getSheetAt(0);
            // 获取gaea框架的模板定义
            Sheet gaeaSheet = wb.getSheetAt(GaeaPoiDefinition.GAEA_DEFINE_SHEET);
            // 获取Gaea WorkBook定义
            result = excelDefineService.getWorkbookDefine(wb);
//            org.gaea.poi.domain.Workbook xmlWorkbook = null;
//            for (Row row : sheet) {
//                // 暂时只在第一行定义导入表达式
//                if (row.getRowNum() == 0) {
//                    Block block = null;// 所有的数据都关联到块，所以必须先有block定义
//                    //遍历row中的所有方格
//                    for (Cell cell : row) {
//                        // 获取Gaea通用导入的Sheet的定义
//                        ExcelSheet excelSheet = excelDefineService.getSheetDefine(cell);
//                        if (excelSheet != null) {
//                            // Workbook的id，就是excelSheet的id，即#GAEA_DEF__BEGIN[]的根id
//                            // 把XML配置的workbook信息合并到当前结果中（相同ID的）
//                            if (StringUtils.isNotEmpty(excelSheet.getId())) {
//                                result.setId(excelSheet.getId());
//                                xmlWorkbook = gaeaPoiXmlConfigParser.getWorkbook(excelSheet.getId());
//                                org.gaea.util.BeanUtils.copyProperties(xmlWorkbook, result);
//                            }
//                            // 获取块并设置块。
//                            if (CollectionUtils.isNotEmpty(excelSheet.getBlockList())) {
//                                // 目前只支持一个sheet
//                                block = excelDefineService.blockParse(result, excelSheet.getBlockList().get(0));
//                            }
//                        }
//                        // 块不为空，数据才能关联到块。块关联到表！
//                        String myExcelRemark = excelDefineService.getCellComment(cell);
//                        if (block != null && StringUtils.isNotEmpty(myExcelRemark)) {
//                            ExcelField excelField = excelDefineService.getField(block, cell, myExcelRemark);
//                            if (block.getFieldDefines() == null) {
//                                block.setFieldDefines(new ArrayList<ExcelField>());
//                            }
//                            block.getFieldDefines().add(excelField);
//                            // 合并Excel定义和XML定义
//                            excelDefineService.combineField(block, excelField);
//                            // 放入 columnIndex：定义的map中。后面获取值需要根据columnIndex获取使用Field
//                            columnFieldDefMap.put(cell.getColumnIndex(), (Field) block.getFieldMap().get(excelField.getName()));
//                            // 获取对应的field定义。同样，本来这里也应该是不同block不一样的。现在暂未实现。
////                            ExcelField excelField = expressParser.parseField(myExcelRemark);
////                            excelField.setColumnIndex(cell.getColumnIndex());
////                            if (block.getFieldDefines() == null) {
////                                block.setFieldDefines(new ArrayList<ExcelField>());
////                            }
////                            block.getFieldDefines().add(excelField);
//                        }
////                    }
//                    }
//                }
//            }
            /* 最后再处理数据。 */
            if (result != null && result.getBlockList() != null && result.getBlockList().size() > 0) {
                Block block = result.getBlockList().get(0);
                // 获取excel数据，放入block。这里本来应该根据block id分别获取的。暂时未实现。
                List<Map<String, String>> blockData = getData(sheet, block.getFieldMap(), gaeaSheet);
                block.setData(blockData);
            }
        } catch (IOException e) {
            logger.error("通用导入excel，解析文件IO错误。", e);
        } catch (InvalidFormatException e) {
            logger.error("通用导入excel，从excel文件输入流创建Workbook失败！", e);
        }
        return result;
    }

    /**
     * 从一个Excel文件的InputStream中读取数据，形成List<Map>的格式返回。
     *
     * @param fileIS
     * @return
     * @throws ValidationFailedException
     */
    public List<Map<String, String>> getData(InputStream fileIS) throws ValidationFailedException {
        List<Map<String, String>> result = null;
        try {
            //根据上述创建的输入流 创建工作簿对象ZA
            Workbook wb = WorkbookFactory.create(fileIS);
            //得到第一页 sheet
            //页Sheet是从0开始索引的。默认先支持一个sheet。
            Sheet sheet = wb.getSheetAt(0);
            // 获取gaea框架的模板定义
            Sheet gaeaSheet = wb.getSheetAt(GaeaPoiDefinition.GAEA_DEFINE_SHEET);
            // 获取数据
            result = getData(sheet, gaeaSheet);
            //关闭输入流
            fileIS.close();
        } catch (IOException e) {
            logger.error("通用导入excel，解析文件IO错误。", e);
        } catch (InvalidFormatException e) {
            logger.error("通用导入excel，从excel文件输入流创建Workbook失败！", e);
        }
        return result;
    }

    /**
     * 从一个Excel文件的InputStream中读取数据，根据fieldDefMap定义获取/转换数据，形成List<Map>的格式返回。
     * <p>
     * 这个主要是方便非模板的方式定义数据。客户端可以灵活自己定义数据格式然后导出。
     * </p>
     *
     * @param fileIS
     * @param fieldDefMap Block的各个字段的gaea定义（包括字段名、数据类型等）。key(cell的gaea定义的name) : value(Field对象)
     * @return
     * @throws ValidationFailedException
     */
    public List<Map<String, String>> getData(InputStream fileIS, Map<String, Field> fieldDefMap) throws ValidationFailedException {
        List<Map<String, String>> result = null;
        try {
            //根据上述创建的输入流 创建工作簿对象ZA
            Workbook wb = WorkbookFactory.create(fileIS);
            //得到第一页 sheet
            //页Sheet是从0开始索引的。默认先支持一个sheet。
            Sheet sheet = wb.getSheetAt(0);
            // 获取gaea框架的模板定义
            Sheet gaeaSheet = wb.getSheetAt(GaeaPoiDefinition.GAEA_DEFINE_SHEET);
            // 获取数据
            result = getData(sheet, fieldDefMap, gaeaSheet);
            //关闭输入流
            fileIS.close();
        } catch (IOException e) {
            logger.error("通用导入excel，解析文件IO错误。", e);
        } catch (InvalidFormatException e) {
            logger.error("通用导入excel，从excel文件输入流创建Workbook失败！", e);
        }
        return result;
    }

    public <T> List<T> getData(InputStream fileIS, Class<T> beanClass) throws ValidationFailedException {
        List<T> results = new ArrayList<T>();
        try {
            //根据上述创建的输入流 创建工作簿对象ZA
            Workbook wb = WorkbookFactory.create(fileIS);
            //得到第一页 sheet
            //页Sheet是从0开始索引的
            Sheet sheet = wb.getSheetAt(0);
            // 获取gaea框架的模板定义
            Sheet gaeaSheet = wb.getSheetAt(GaeaPoiDefinition.GAEA_DEFINE_SHEET);
            // 获取数据
            List<Map<String, String>> dataList = getData(sheet, gaeaSheet);
//            for (int i = 0; dataList != null && i < dataList.size(); i++) {
//                // TODO 这里可以优化！没必要放在循环里面啊！！
//                T bean = BeanUtils.instantiate(beanClass);
//                BeanWrapper wrapper = PropertyAccessorFactory.forBeanPropertyAccess(bean);
//                wrapper.setAutoGrowNestedPaths(true);
//                Map<String, String> dataMap = dataList.get(i);
//                wrapper.setPropertyValues(dataMap);
//                results.add(bean);
//            }
            results = org.gaea.util.BeanUtils.getData(dataList, beanClass);
        } catch (IOException e) {
            logger.error("通用导入excel，解析文件IO错误。", e);
        } catch (InvalidFormatException e) {
            logger.error("通用导入excel，从excel文件输入流创建Workbook失败！", e);
        }
        return results;
    }

    /**
     * 从一个Sheet对象中，解析出每一行的数据，并组成List返回。
     * <p>每一个Map都是一行数据，key=列名 value=值</p>
     *
     * @param sheet
     * @param gaeaSheet gaea导入模板定义的专有sheet
     * @return
     * @throws ValidationFailedException
     */
    public List<Map<String, String>> getData(Sheet sheet, Sheet gaeaSheet) throws ValidationFailedException {
        return getData(sheet, null, gaeaSheet);
    }

    /**
     * 从一个Sheet对象中，解析出每一行的数据，并组成List返回。
     * <p>每一个Map都是一行数据，key=列名 value=值</p>
     *
     * @param sheet
     * @param fieldDefMap Block的各个字段的gaea定义（包括字段名、数据类型等）。key(cell的gaea定义的name) : value(Field对象)
     * @param gaeaSheet   gaea导入模板定义的专有sheet
     * @return
     * @throws ValidationFailedException
     */
    public List<Map<String, String>> getData(Sheet sheet, Map<String, Field> fieldDefMap, Sheet gaeaSheet) throws ValidationFailedException {
        List<Map<String, String>> results = new ArrayList<Map<String, String>>();
        Map<Integer, Field> columnDefMap = null;
        // 获取导入模板定义
        if (gaeaSheet == null) {
            throw new ValidationFailedException("导入Excel缺少Gaea框架专有的模板定义sheet。无法执行导入读取操作。");
        }
        columnDefMap = excelDefineService.getFieldsDefine(gaeaSheet, fieldDefMap);// 获取Excel中列定义，合并外部传入定义（如果有），并转换为基于列索引的map
        //利用foreach循环 遍历sheet中的所有行
        for (Row row : sheet) {
            Map<String, String> rowValueMap = new HashMap<String, String>();
            if (row.getRowNum() == GaeaPoiDefinition.EXCEL_TITLE_ROW) {
                // do nothing
                // 第一行就是一般的Excel表普通title(表示会略过, 不会当数据读取)
            } else {
                //遍历row中的所有方格
                for (Cell cell : row) {
//                if (cell.getRowIndex() == 0) {
//                    // 获取Gaea通用导入的字段定义
//                    ExcelField fieldDef = getFieldDefine(cell);
//                    if (fieldDef != null) {
//                        columnDefMap.put(cell.getColumnIndex(), fieldDef);
//                    }
//                }
                    // 存在Gaea导入定义的列，才取值放入对象中
                    if (columnDefMap.get(cell.getColumnIndex()) != null) {
//                    String dataType = "";
//                    if(columnIndexDefMap!=null){
                        String dataType = columnDefMap.get(cell.getColumnIndex()).getDataType();
//                    }
                        // 获取Excel单元格的值。统一String类型。
                        String value = GaeaPoiUtils.getCellStringValue(cell, dataType);
                        String key = columnDefMap.get(cell.getColumnIndex()).getName();
                        rowValueMap.put(key, value);
                    }
                }
                results.add(rowValueMap);
            }
        }
        return results;
    }

//    /**
//     * 解析Sheet对象，返回Gaea POI的表达式定义的各个块。
//     *
//     * @param sheet
//     * @return
//     * @throws ValidationFailedException
//     */
//    public List<Block> getFieldDefines(Sheet sheet) throws ValidationFailedException {
//        List<Block> blockList = new ArrayList<Block>();
//        //利用foreach循环 遍历sheet中的所有行
//        for (Row row : sheet) {
//            Block block = new Block();
//            block.setFieldDefines(new ArrayList<ExcelField>());
//            //遍历row中的所有方格
//            for (Cell cell : row) {
//                // 在excel第一行中找定义
//                if (cell.getRowIndex() == 0) {
//                    String myExcelRemark = cell.getCellComment().getString().getString();
//                    ExcelField fieldDef = expressParser.parseField(myExcelRemark);
//                    fieldDef.setColumnIndex(cell.getColumnIndex());
//                    block.getFieldDefines().add(fieldDef);
//                }
//            }
//            blockList.add(block);
//        }
//        return blockList;
//    }

    public void readToDB(InputStream excelIS) {

    }

//    private Block blockParse(org.gaea.poi.domain.Workbook workbook, ExcelBlock excelBlock) {
//        Block block = null;
//        // 合并XML的块配置和Excel文件的块配置信息。
//        if (workbook.getBlockList() == null) {
//            workbook.setBlockList(new ArrayList<Block>());
//        }
//        if (workbook.getBlockList() != null) {
//            if (workbook.getBlockList().isEmpty()) {
//                block = new Block();
//                workbook.getBlockList().add(block);
//            } else {
//                block = workbook.getBlockList().get(0);
//            }
//            org.gaea.util.BeanUtils.copyProperties(excelBlock, block);
//        }
//        return block;
//    }

//    private void combineField(Block block, ExcelField excelField) {
//        boolean hasXmlFieldDef = false;
//        Map<String, Field> fieldMap = block.getFieldMap();
//        Field f = fieldMap.get(excelField.getName());// 获取是否有对应的XML定义的field
//        // 如果XML定义和Excel定义同时存在，Excel定义相关值覆盖XML定义（但XML是父级，一些关键值EXCEL是覆盖不了的）
//        if (f == null) {
//            f = new Field();
//        }
//        org.gaea.util.BeanUtils.copyProperties(excelField, f);
//        // 如果没有找到XML定义和Excel定义匹配的，创建一个，并放入block中
//        if (!hasXmlFieldDef) {
//            block.getFieldMap().put(excelField.getName(), f);
//        }
//    }

//    private ExcelField getField(Block block, Cell cell, String myExcelRemark) throws ValidationFailedException {
//        // 获取对应的field定义。同样，本来这里也应该是不同block不一样的。现在暂未实现。
//        ExcelField excelField = expressParser.parseField(myExcelRemark);
//        excelField.setColumnIndex(cell.getColumnIndex());
//        return excelField;
//    }

    /**
     * 获取excel里一个cell的值。根据类型判断，但最后统一转换为String。
     * 因为POI的接口没有统一返回cell中的值的。
     *
     * @param cell
     * @param dataType XML定义的读取类型。为空则按Excel单元格类型转换。
     * @return
     */
//    private String getCellStringValue(Cell cell, String dataType) {
//        String value = "";
//        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
//            value = cell.getStringCellValue();
//        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
//            // 如果是日期类型
//            if (HSSFDateUtil.isCellDateFormatted(cell)) {
//                value = cell.getDateCellValue().toString();
//                // 如果XML配置了对应的日期类型，则按XML配置转换；否则按系统配置的默认年月日时分秒转换。
//                if (Field.DATA_TYPE_DATE.equalsIgnoreCase(dataType)) {
//                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATE_FORMAT));
//                } else if (Field.DATA_TYPE_TIME.equalsIgnoreCase(dataType)) {
//                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_TIME_FORMAT));
//                } else if (Field.DATA_TYPE_DATETIME.equalsIgnoreCase(dataType)) {
//                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT));
//                } else {
//                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT));
//                }
//                return value;
//            }
//            value = String.valueOf(cell.getNumericCellValue());
//        }
//        return value;
//    }

//    private ExcelField getFieldDefine(Cell cell) throws ValidationFailedException {
//        String myExcelRemark = null;
//        ExcelField fieldDef = null;
//        if (cell.getCellComment() != null) {
//            myExcelRemark = cell.getCellComment().getString().getString();
//            fieldDef = expressParser.parseField(myExcelRemark);
//        }
//        return fieldDef;
//    }

//    private ExcelSheet getSheetDefine(Cell cell) throws ValidationFailedException {
//        String myExcelRemark = null;
//        ExcelSheet excelSheet = null;
//        if (cell.getCellComment() != null) {
//            myExcelRemark = cell.getCellComment().getString().getString();
//            excelSheet = expressParser.parseSheet(myExcelRemark);
//        }
//        return excelSheet;
//    }

//    private String getCellComment(Cell cell) {
//        String excelRemark = null;
//        if (cell.getCellComment() != null) {
//            excelRemark = cell.getCellComment().getString().getString();
//        }
//        return excelRemark;
//    }
}
