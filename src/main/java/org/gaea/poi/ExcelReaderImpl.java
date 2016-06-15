package org.gaea.poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.config.GaeaPoiDefinition;
import org.gaea.poi.domain.*;
import org.gaea.poi.util.ExpressParser;
import org.gaea.poi.xml.GaeaPoiXmlConfigParser;
import org.gaea.util.GaeaPropertiesReader;
import org.omg.CORBA.PRIVATE_MEMBER;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeanWrapper;
import org.springframework.beans.PropertyAccessorFactory;
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

    public org.gaea.poi.domain.Workbook getWorkbook(InputStream fileIS) throws ValidationFailedException {
        org.gaea.poi.domain.Workbook result = new org.gaea.poi.domain.Workbook();
        result.setBlockList(new ArrayList<Block>());
        Map<Integer, Field> columnFieldDefMap = new HashMap<Integer, Field>();
        try {
            //根据上述创建的输入流 创建工作簿对象ZA
            Workbook wb = WorkbookFactory.create(fileIS);
            //得到第一页 sheet
            //页Sheet是从0开始索引的。默认先支持一个sheet。
            Sheet sheet = wb.getSheetAt(0);
            org.gaea.poi.domain.Workbook xmlWorkbook = null;
            for (Row row : sheet) {
                // 暂时只在第一行定义导入表达式
                if (row.getRowNum() == 0) {
                    Block block = null;// 所有的数据都关联到块，所以必须先有block定义
                    //遍历row中的所有方格
                    for (Cell cell : row) {
//                    if (cell.getRowIndex() == 0) {
                        // 获取Gaea通用导入的Sheet的定义
                        ExcelSheet excelSheet = getSheetDefine(cell);
//                        if(cell.getCellComment()!=null) {
//                            myExcelRemark = cell.getCellComment().getString().getString();
//                            excelSheet = expressParser.parseSheet(myExcelRemark);
//                        }
                        if (excelSheet != null) {
                            // Workbook的id，就是excelSheet的id，即#GAEA_DEF__BEGIN[]的根id
                            // 把XML配置的workbook信息合并到当前结果中（相同ID的）
                            if (StringUtils.isNotEmpty(excelSheet.getId())) {
                                result.setId(excelSheet.getId());
                                xmlWorkbook = gaeaPoiXmlConfigParser.getWorkbook(excelSheet.getId());
                                org.gaea.util.BeanUtils.copyProperties(xmlWorkbook, result);
                            }
                            // 获取块并设置块。
                            if (excelSheet.getBlock() != null) {
                                block = blockParse(result, excelSheet.getBlock());
                            }
                        }
                        // 块不为空，数据才能关联到块。块关联到表！
                        String myExcelRemark = getCellComment(cell);
                        if (block != null && StringUtils.isNotEmpty(myExcelRemark)) {
                            ExcelField excelField = getField(block, cell, myExcelRemark);
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
            /* 最后再处理数据。 */
            if (result != null && result.getBlockList() != null && result.getBlockList().size() > 0) {
                Block block = result.getBlockList().get(0);
                // 获取excel数据，放入block。这里本来应该根据block id分别获取的。暂时未实现。
                List<Map<String, String>> blockData = getData(sheet,columnFieldDefMap);
                block.setData(blockData);
            }
            //关闭输入流
//            fileIS.close();
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
            result = getData(sheet);
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
//        Map<String, String> excelRowValue = new HashMap<String, String>();
//        Map<Integer, ExcelField> columnDefMap = new HashMap<Integer, ExcelField>();
        List<T> results = new ArrayList<T>();
        try {
            //根据上述创建的输入流 创建工作簿对象ZA
            Workbook wb = WorkbookFactory.create(fileIS);
            //得到第一页 sheet
            //页Sheet是从0开始索引的
            Sheet sheet = wb.getSheetAt(0);
            List<Map<String, String>> dataList = getData(sheet);
            for (int i = 0; dataList != null && i < dataList.size(); i++) {
                T bean = BeanUtils.instantiate(beanClass);
                BeanWrapper wrapper = PropertyAccessorFactory.forBeanPropertyAccess(bean);
                wrapper.setAutoGrowNestedPaths(true);
                Map<String, String> dataMap = dataList.get(i);
                wrapper.setPropertyValues(dataMap);
                results.add(bean);
            }
            //关闭输入流
//            fileIS.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return results;
    }

    /**
     * 从一个Sheet对象中，解析出每一行的数据，并组成List返回。
     * <p>每一个Map都是一行数据，key=列名 value=值</p>
     *
     * @param sheet
     * @return
     * @throws ValidationFailedException
     */
    public List<Map<String, String>> getData(Sheet sheet) throws ValidationFailedException {
        return getData(sheet, null);
    }

    /**
     * 从一个Sheet对象中，解析出每一行的数据，并组成List返回。
     * <p>每一个Map都是一行数据，key=列名 value=值</p>
     *
     * @param sheet
     * @param columnIndexDefMap    key: cell的columnIndex,Excel的列下标 value：Field对象
     * @return
     * @throws ValidationFailedException
     */
    public List<Map<String, String>> getData(Sheet sheet,Map<Integer,Field> columnIndexDefMap) throws ValidationFailedException {
        List<Map<String, String>> results = new ArrayList<Map<String, String>>();
        Map<Integer, ExcelField> columnDefMap = new HashMap<Integer, ExcelField>();
        //利用foreach循环 遍历sheet中的所有行
        for (Row row : sheet) {
            Map<String, String> rowValueMap = new HashMap<String, String>();
            //遍历row中的所有方格
            for (Cell cell : row) {
                if (cell.getRowIndex() == 0) {
                    // 获取Gaea通用导入的字段定义
                    ExcelField fieldDef = getFieldDefine(cell);
//                    String myExcelRemark = null;
//                    if(cell.getCellComment()!=null) {
//                        myExcelRemark = cell.getCellComment().getString().getString();
//                        ExcelField fieldDef = expressParser.parseField(myExcelRemark);
                    if (fieldDef != null) {
                        columnDefMap.put(cell.getColumnIndex(), fieldDef);
                    }
//                    }
                }
                // 存在Gaea导入定义的列，才取值放入对象中
                if (columnDefMap.get(cell.getColumnIndex()) != null) {
                    String readType = "";
                    if(columnIndexDefMap!=null){
                        readType = columnIndexDefMap.get(cell.getColumnIndex()).getReadType();
                    }
                    // 获取Excel单元格的值。统一String类型。
                    String value = getCellStringValue(cell,readType);
                    String key = columnDefMap.get(cell.getColumnIndex()).getName();
                    rowValueMap.put(key, value);
                }
            }
            results.add(rowValueMap);
        }
        return results;
    }

    /**
     * 解析Sheet对象，返回Gaea POI的表达式定义的各个块。
     *
     * @param sheet
     * @return
     * @throws ValidationFailedException
     */
    public List<Block> getFieldDefines(Sheet sheet) throws ValidationFailedException {
        List<Block> blockList = new ArrayList<Block>();
        //利用foreach循环 遍历sheet中的所有行
        for (Row row : sheet) {
            Block block = new Block();
            block.setFieldDefines(new ArrayList<ExcelField>());
            //遍历row中的所有方格
            for (Cell cell : row) {
                // 在excel第一行中找定义
                if (cell.getRowIndex() == 0) {
                    String myExcelRemark = cell.getCellComment().getString().getString();
                    ExcelField fieldDef = expressParser.parseField(myExcelRemark);
                    fieldDef.setColumnIndex(cell.getColumnIndex());
                    block.getFieldDefines().add(fieldDef);
                }
            }
            blockList.add(block);
        }
        return blockList;
    }

//    public <T> GaeaPoiResultGroup<T> getDataTest(InputStream fileIS, Class<T> beanClass) throws ValidationFailedException {
//        GaeaPoiResultGroup<T> resultGroup = new GaeaPoiResultGroup<T>();
//        GaeaPoiResultSet<T> resultSet = new GaeaPoiResultSet<T>();
//        T bean = BeanUtils.instantiate(beanClass);
//        BeanWrapper wrapper = PropertyAccessorFactory.forBeanPropertyAccess(bean);
//        wrapper.setAutoGrowNestedPaths(true);
//        Map<String, String> excelRowValue = new HashMap<String, String>();
//        Map<Integer, ExcelField> columnDefMap = new HashMap<Integer, ExcelField>();
//        List<T> results = new ArrayList<T>();
//        try {
//            //根据上述创建的输入流 创建工作簿对象ZA
//            Workbook wb = WorkbookFactory.create(fileIS);
//            //得到第一页 sheet
//            //页Sheet是从0开始索引的
//            Sheet sheet = wb.getSheetAt(0);
//            //利用foreach循环 遍历sheet中的所有行
//            for (Row row : sheet) {
//                //遍历row中的所有方格
//                for (Cell cell : row) {
//                    if(cell.getCellComment()==null){
//                        continue;
//                    }
//                    if (cell.getRowIndex() == 0) {
//                        String myExcelRemark = cell.getCellComment().getString().getString();
//                        ExcelField fieldDef = expressParser.parseField(myExcelRemark);
//                        columnDefMap.put(cell.getColumnIndex(), fieldDef);
//                    }
//                    String value = getCellStringValue(cell);
//                    String key = columnDefMap.get(cell.getColumnIndex()).getName();
//                    excelRowValue.put(key, value);
////                    String stringTemplate = "cellComment:{0} cellType:{1} columnIndex:{2} rowIndex:{3}";
////                    String log = MessageFormat.format(stringTemplate,cell.getCellComment().getString().getString(),cell.getCellType(),cell.getColumnIndex(),cell.getRowIndex());
//                    //输出方格中的内容，以空格间隔
////                    System.out.print(log);
//                    System.out.print(cell.toString() + "  ");
//
//                }
//                //每一个行输出之后换行
//                System.out.println();
//
//                wrapper.setPropertyValues(excelRowValue);
//                results.add(bean);
//            }
//            resultSet.setData(results);
//            resultGroup.setResult(resultSet);
//            //关闭输入流
//            fileIS.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        } catch (InvalidFormatException e) {
//            e.printStackTrace();
//        }
//        return resultGroup;
//    }

    public void readToDB(InputStream excelIS) {

    }

    private Block blockParse(org.gaea.poi.domain.Workbook workbook, ExcelBlock excelBlock) {
        Block block = null;
        // 合并XML的块配置和Excel文件的块配置信息。
        if (workbook.getBlockList() == null) {
            workbook.setBlockList(new ArrayList<Block>());
        }
        if (workbook.getBlockList() != null) {
            if (workbook.getBlockList().isEmpty()) {
                block = new Block();
                workbook.getBlockList().add(block);
            } else {
                block = workbook.getBlockList().get(0);
            }
            org.gaea.util.BeanUtils.copyProperties(excelBlock, block);
        }
        return block;
    }

    private void combineField(Block block, ExcelField excelField) {
        boolean hasXmlFieldDef = false;
        Map<String,Field> fieldMap = block.getFieldMap();
        Field f = fieldMap.get(excelField.getName());// 获取是否有对应的XML定义的field
        // 如果XML定义和Excel定义同时存在，Excel定义相关值覆盖XML定义（但XML是父级，一些关键值EXCEL是覆盖不了的）
        if(f==null){
            f = new Field();
        }
//            if (f.getName().equalsIgnoreCase(excelField.getName())) {
                org.gaea.util.BeanUtils.copyProperties(excelField, f);
//                hasCombine = true;
//                break;
//            }
        // 如果没有找到XML定义和Excel定义匹配的，创建一个，并放入block中
//        if (!hasCombine) {
//            Field field = new Field();
//            org.gaea.util.BeanUtils.copyProperties(excelField, field);
        if(!hasXmlFieldDef) {
            block.getFieldMap().put(excelField.getName(),f);
        }
//        }
    }

    private ExcelField getField(Block block, Cell cell, String myExcelRemark) throws ValidationFailedException {
        // 获取对应的field定义。同样，本来这里也应该是不同block不一样的。现在暂未实现。
        ExcelField excelField = expressParser.parseField(myExcelRemark);
        excelField.setColumnIndex(cell.getColumnIndex());
//        if (block.getFieldDefines() == null) {
//            block.setFieldDefines(new ArrayList<ExcelField>());
//        }
//        block.getFieldDefines().add(excelField);
        return excelField;
    }

    /**
     * 获取excel里一个cell的值。根据类型判断，但最后统一转换为String。
     * 因为POI的接口没有统一返回cell中的值的。
     *
     * @param cell
     * @param readType XML定义的读取类型。为空则按Excel单元格类型转换。
     * @return
     */
    private String getCellStringValue(Cell cell,String readType) {
        String value = "";
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            value = cell.getStringCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            // 如果是日期类型
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                value = cell.getDateCellValue().toString();
                // 如果XML配置了对应的日期类型，则按XML配置转换；否则按系统配置的默认年月日时分秒转换。
                if(Field.READ_TYPE_DATE.equalsIgnoreCase(readType)){
                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATE_FORMAT));
                }else if(Field.READ_TYPE_TIME.equalsIgnoreCase(readType)){
                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_TIME_FORMAT));
                }else if(Field.READ_TYPE_DATETIME.equalsIgnoreCase(readType)){
                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT));
                }else{
                    value = DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT));
                }
//                System.out.println(DateFormatUtils.format(cell.getDateCellValue().getTime(), cacheProperties.get(GaeaPoiDefinition.POI_DEFAULT_DATETIME_FORMAT)));
                return value;
            }
            value = String.valueOf(cell.getNumericCellValue());
        }
        return value;
    }

    private ExcelField getFieldDefine(Cell cell) throws ValidationFailedException {
        String myExcelRemark = null;
        ExcelField fieldDef = null;
        if (cell.getCellComment() != null) {
            myExcelRemark = cell.getCellComment().getString().getString();
            fieldDef = expressParser.parseField(myExcelRemark);
        }
        return fieldDef;
    }

    private ExcelSheet getSheetDefine(Cell cell) throws ValidationFailedException {
        String myExcelRemark = null;
        ExcelSheet excelSheet = null;
        if (cell.getCellComment() != null) {
            myExcelRemark = cell.getCellComment().getString().getString();
            excelSheet = expressParser.parseSheet(myExcelRemark);
        }
        return excelSheet;
    }

    private String getCellComment(Cell cell) {
        String excelRemark = null;
        if (cell.getCellComment() != null) {
            excelRemark = cell.getCellComment().getString().getString();
        }
        return excelRemark;
    }
}
