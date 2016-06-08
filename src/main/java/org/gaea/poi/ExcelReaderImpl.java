package org.gaea.poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.domain.*;
import org.gaea.poi.util.ExpressParser;
import org.gaea.poi.xml.GaeaPoiXmlConfigParser;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeanWrapper;
import org.springframework.beans.PropertyAccessorFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

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

    public org.gaea.poi.domain.Workbook getWorkbook(InputStream fileIS) throws ValidationFailedException {
        org.gaea.poi.domain.Workbook result = new org.gaea.poi.domain.Workbook();
        result.setBlockList(new ArrayList<Block>());
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
                                block = blockParse(result,excelSheet.getBlock());
                            }
                        }
                        // 块不为空，数据才能关联到块。块关联到表！
                        String myExcelRemark = getCellComment(cell);
                        if (block != null && StringUtils.isNotEmpty(myExcelRemark)) {
                            parseField(block,cell,myExcelRemark);
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
                List<Map<String, String>> blockData = getData(sheet);
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
        GaeaPoiResultGroup<T> resultGroup = new GaeaPoiResultGroup<T>();
        GaeaPoiResultSet<T> resultSet = new GaeaPoiResultSet<T>();
        T bean = BeanUtils.instantiate(beanClass);
        BeanWrapper wrapper = PropertyAccessorFactory.forBeanPropertyAccess(bean);
        wrapper.setAutoGrowNestedPaths(true);
        Map<String, String> excelRowValue = new HashMap<String, String>();
        Map<Integer, ExcelField> columnDefMap = new HashMap<Integer, ExcelField>();
        List<T> results = new ArrayList<T>();
        try {
            //根据上述创建的输入流 创建工作簿对象ZA
            Workbook wb = WorkbookFactory.create(fileIS);
            //得到第一页 sheet
            //页Sheet是从0开始索引的
            Sheet sheet = wb.getSheetAt(0);
            List<Map<String, String>> dataList = getData(sheet);
            for (int i = 0; dataList != null && i < dataList.size(); i++) {
                Map<String, String> dataMap = dataList.get(i);
                wrapper.setPropertyValues(dataMap);
                results.add(bean);
            }
            //利用foreach循环 遍历sheet中的所有行
            resultSet.setData(results);
            resultGroup.setResult(resultSet);
            //关闭输入流
            fileIS.close();
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
        List<Map<String, String>> results = new ArrayList<Map<String, String>>();
        Map<String, String> excelRowValue = new HashMap<String, String>();
        Map<Integer, ExcelField> columnDefMap = new HashMap<Integer, ExcelField>();
        //利用foreach循环 遍历sheet中的所有行
        for (Row row : sheet) {
            //遍历row中的所有方格
            for (Cell cell : row) {
                if (cell.getRowIndex() == 0) {
                    // 获取Gaea通用导入的字段定义
                    ExcelField fieldDef = getFieldDefine(cell);
//                    String myExcelRemark = null;
//                    if(cell.getCellComment()!=null) {
//                        myExcelRemark = cell.getCellComment().getString().getString();
//                        ExcelField fieldDef = expressParser.parseField(myExcelRemark);
                    if(fieldDef!=null) {
                        columnDefMap.put(cell.getColumnIndex(), fieldDef);
                    }
//                    }
                }
                // 存在Gaea导入定义的列，才取值放入对象中
                if(columnDefMap.get(cell.getColumnIndex())!=null) {
                    // 获取Excel单元格的值。统一String类型。
                    String value = getCellStringValue(cell);
                    String key = columnDefMap.get(cell.getColumnIndex()).getName();
                    excelRowValue.put(key, value);
                }
            }
            results.add(excelRowValue);
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

    public <T> GaeaPoiResultGroup<T> getDataTest(InputStream fileIS, Class<T> beanClass) throws ValidationFailedException {
        GaeaPoiResultGroup<T> resultGroup = new GaeaPoiResultGroup<T>();
        GaeaPoiResultSet<T> resultSet = new GaeaPoiResultSet<T>();
        T bean = BeanUtils.instantiate(beanClass);
        BeanWrapper wrapper = PropertyAccessorFactory.forBeanPropertyAccess(bean);
        wrapper.setAutoGrowNestedPaths(true);
        Map<String, String> excelRowValue = new HashMap<String, String>();
        Map<Integer, ExcelField> columnDefMap = new HashMap<Integer, ExcelField>();
        List<T> results = new ArrayList<T>();
        try {
            //根据上述创建的输入流 创建工作簿对象ZA
            Workbook wb = WorkbookFactory.create(fileIS);
            //得到第一页 sheet
            //页Sheet是从0开始索引的
            Sheet sheet = wb.getSheetAt(0);
            //利用foreach循环 遍历sheet中的所有行
            for (Row row : sheet) {
                //遍历row中的所有方格
                for (Cell cell : row) {
                    if(cell.getCellComment()==null){
                        continue;
                    }
                    if (cell.getRowIndex() == 0) {
                        String myExcelRemark = cell.getCellComment().getString().getString();
                        ExcelField fieldDef = expressParser.parseField(myExcelRemark);
                        columnDefMap.put(cell.getColumnIndex(), fieldDef);
                    }
                    String value = getCellStringValue(cell);
                    String key = columnDefMap.get(cell.getColumnIndex()).getName();
                    excelRowValue.put(key, value);
//                    String stringTemplate = "cellComment:{0} cellType:{1} columnIndex:{2} rowIndex:{3}";
//                    String log = MessageFormat.format(stringTemplate,cell.getCellComment().getString().getString(),cell.getCellType(),cell.getColumnIndex(),cell.getRowIndex());
                    //输出方格中的内容，以空格间隔
//                    System.out.print(log);
                    System.out.print(cell.toString() + "  ");

                }
                //每一个行输出之后换行
                System.out.println();

                wrapper.setPropertyValues(excelRowValue);
                results.add(bean);
            }
            resultSet.setData(results);
            resultGroup.setResult(resultSet);
            //关闭输入流
            fileIS.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return resultGroup;
    }

    public void readToDB(InputStream excelIS) {

    }

    private Block blockParse(org.gaea.poi.domain.Workbook workbook, ExcelBlock excelBlock){
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
//                                    if (block == null) {
//                                        block = new Block();
//                                    }
            org.gaea.util.BeanUtils.copyProperties(excelBlock, block);
        }
        return block;
    }

    private void parseField(Block block, Cell cell, String myExcelRemark) throws ValidationFailedException {
        // 获取对应的field定义。同样，本来这里也应该是不同block不一样的。现在暂未实现。
        ExcelField excelField = expressParser.parseField(myExcelRemark);
        excelField.setColumnIndex(cell.getColumnIndex());
        if (block.getFieldDefines() == null) {
            block.setFieldDefines(new ArrayList<ExcelField>());
        }
        block.getFieldDefines().add(excelField);
    }

    /**
     * 获取excel里一个cell的值。根据类型判断，但最后统一转换为String。
     * 因为POI的接口没有统一返回cell中的值的。
     *
     * @param cell
     * @return
     */
    private String getCellStringValue(Cell cell) {
        String value = "";
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            value = cell.getStringCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                value = cell.getDateCellValue().toString();
                System.out.println(DateFormatUtils.format(cell.getDateCellValue().getTime(), "yyyy-MM-dd HH:mm:ss"));
                return value;
            }
            value = String.valueOf(cell.getNumericCellValue());
        }
        return value;
    }

    private ExcelField getFieldDefine(Cell cell) throws ValidationFailedException {
        String myExcelRemark = null;
        ExcelField fieldDef = null;
        if(cell.getCellComment()!=null) {
            myExcelRemark = cell.getCellComment().getString().getString();
            fieldDef = expressParser.parseField(myExcelRemark);
        }
        return fieldDef;
    }

    private ExcelSheet getSheetDefine(Cell cell) throws ValidationFailedException {
        String myExcelRemark = null;
        ExcelSheet excelSheet = null;
        if(cell.getCellComment()!=null) {
            myExcelRemark = cell.getCellComment().getString().getString();
            excelSheet = expressParser.parseSheet(myExcelRemark);
        }
        return excelSheet;
    }

    private String getCellComment(Cell cell){
        String excelRemark = null;
        if(cell.getCellComment()!=null) {
            excelRemark = cell.getCellComment().getString().getString();
        }
        return excelRemark;
    }
}
