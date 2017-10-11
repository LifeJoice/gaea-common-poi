package org.gaea.poi;

import com.fasterxml.jackson.core.JsonProcessingException;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.gaea.data.dataset.Item;
import org.gaea.exception.ProcessFailedException;
import org.gaea.exception.SysInitException;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.cache.GaeaPoiCache;
import org.gaea.poi.domain.Block;
import org.gaea.poi.domain.ExcelSheet;
import org.gaea.poi.domain.ExcelTemplate;
import org.gaea.poi.domain.Field;
import org.gaea.poi.export.ExcelExport;
import org.gaea.poi.util.ExpressParser;
import org.gaea.poi.util.GaeaPoiUtils;
import org.joda.time.DateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * excel导出的核心类。
 * Created by iverson on 2016/10/26.
 */
@Component
public class ExcelExportImpl implements ExcelExport {
    private final Logger logger = LoggerFactory.getLogger(ExcelExportImpl.class);
    // 单个单元格的默认宽度
    public static final short DEFAULT_CELL_WIDTH = 15;
    public static final short DEFAULT_FLUSH_WORKBOOK_NUM = 1000;// 默认每1000条就写一次文件
    @Autowired
    private GaeaPoiCache gaeaPoiCache;

    /**
     * 导出excel文件。
     *
     * @param excelTemplateId 导出的gaea excel模板id。一般会从缓存中读取。
     * @param data            数据。可以为空。空的话，只导出列头。
     * @param fileDir         excel文件目录。一般是本地临时目录。最后返回的是file。
     * @return
     * @throws SysInitException
     * @throws ValidationFailedException
     * @throws ProcessFailedException
     */
    public File export(String excelTemplateId, List<? extends Map> data, String fileDir) throws SysInitException, ValidationFailedException, ProcessFailedException {
        if (StringUtils.isEmpty(excelTemplateId)) {
            throw new IllegalArgumentException("导出的Excel Template（模板）ID不允许为空！");
        }
        if (StringUtils.isEmpty(fileDir)) {
            throw new IllegalArgumentException("导出Excel，文件目录位置不允许为空！需要用于本地缓存文件。");
        }
        ExcelTemplate excelTemplate = gaeaPoiCache.getExcelTemplate(excelTemplateId);
        if (excelTemplate == null) {
            throw new ValidationFailedException("导出失败！通过id找到的对应excel template为空！");
        }

        org.gaea.poi.domain.Sheet gaeaSheet = null;
        Block gaeaBlock = null;
        Map<String, Field> fieldMap = null;
        /**
         * 获取field map（即列定义）。主要用于生成列头。
         */
        if (CollectionUtils.isNotEmpty(excelTemplate.getExcelSheetList())) {
            // 当前默认只支持一个sheet
            gaeaSheet = excelTemplate.getExcelSheetList().get(0);
            if (gaeaSheet != null && CollectionUtils.isNotEmpty(gaeaSheet.getBlockList())) {
                // 当前只支持一个block
                gaeaBlock = gaeaSheet.getBlockList().get(0);
                fieldMap = gaeaBlock.getFieldMap();
            }
        }
        if (MapUtils.isEmpty(fieldMap)) {
            throw new ValidationFailedException("excel template的模板定义，缺失Field的定义！");
        }
        // 从模板，获取ExcelSheet定义对象
        ExcelSheet excelSheet = ExpressParser.createSheet(excelTemplateId);
        return export(data, null, fieldMap, excelTemplate.getFileName(), fileDir, excelSheet);
    }

    /**
     * @param data      可以为空。空只导出标题行。可作模板下载用。
     * @param sheetName
     * @param fieldMap  列定义。
     * @param fileName  文件名。不需要带后缀。带了一般也没错。
     * @param fileDir   文件存放目录。
     * @return
     * @throws ValidationFailedException
     */
    public File export(List<? extends Map> data, String sheetName, Map<String, Field> fieldMap, String fileName, String fileDir) throws ValidationFailedException, ProcessFailedException {
        ExcelSheet excelSheet = new ExcelSheet();
        excelSheet.setId("AUTO");
        return export(data, sheetName, fieldMap, fileName, fileDir, excelSheet);
    }

    /**
     * 如果excelSheet对象不为空，导出的Excel即为导入Excel模板！可以直接用回导出文件再导入。
     *
     * @param data       可以为空。空只导出标题行。可作模板下载用。
     * @param sheetName
     * @param fieldMap   列定义。
     * @param fileName   文件名。不需要带后缀。带了一般也没错。
     * @param fileDir    文件存放目录。
     * @param excelSheet 可以为空！对应的gaea excel sheet定义对象。如果有，可以导出gaea的模板定义。
     * @return
     * @throws ValidationFailedException
     * @throws ProcessFailedException
     */
    private File export(List<? extends Map> data, String sheetName, Map<String, Field> fieldMap, String fileName, String fileDir, ExcelSheet excelSheet) throws ValidationFailedException, ProcessFailedException {
        if (MapUtils.isEmpty(fieldMap)) {
            throw new ValidationFailedException("excel template的模板定义，缺失Field的定义！");
        }
        if (StringUtils.isEmpty(fileDir)) {
            throw new IllegalArgumentException("导出Excel，文件目录位置不允许为空！需要用于本地缓存文件。");
        }
        // 声明一个工作薄
        // 大于1000行时会把之前的行写入硬盘
        SXSSFWorkbook workbook = new SXSSFWorkbook(DEFAULT_FLUSH_WORKBOOK_NUM);
        // 生成一个表格
        SXSSFSheet sheet = null;
        if (StringUtils.isEmpty(sheetName)) {
            sheet = workbook.createSheet();
        } else {
            sheet = workbook.createSheet(sheetName);
        }
        // 设置表格默认列宽度
        sheet.setDefaultColumnWidth(DEFAULT_CELL_WIDTH);

        // 声明一个画图的顶级管理器
        CreationHelper factory = workbook.getCreationHelper();
        Drawing drawing = sheet.createDrawingPatriarch();

        // When the comment box is visible, have it show in a 1x3 space
        ClientAnchor anchor = factory.createClientAnchor();

        int rowIndex = 0;
        // 第一行（隐藏），Gaea框架的模板定义
        if (excelSheet != null) {
            SXSSFRow firstRow = sheet.createRow(rowIndex);
            createGaeaRow(excelSheet, firstRow, fieldMap, drawing, anchor, factory);
            rowIndex++;
        }

        // 第二行，先写标题
        SXSSFRow row = sheet.createRow(rowIndex);
        createTitleRow(sheet, row, fieldMap, drawing, anchor, factory);
        String[] fieldKeys = fieldMap.keySet().toArray(new String[]{});
//        for (int j = 0; j < fieldKeys.length; j++) {
//            String fieldKey = fieldKeys[j];
//            /**
//             * 用结果集data的map的key，去查找title map的key。这两者应该是一致的。
//             * 参考CommonViewQueryServiceImpl.query，是经过schemaDataService.transformViewData处理过的。
//             */
//            Field titleField = fieldMap.get(fieldKey);
//            String colTitle = titleField == null ? "" : titleField.getTitle();
//            SXSSFCell cell = row.createCell(j);
//            cell.setCellValue(colTitle);
//            // 设定列宽
//            if (NumberUtils.isNumber(titleField.getWidth())) {
//                sheet.setColumnWidth(j, Integer.parseInt(titleField.getWidth()) * 256); // 根据API，这里设的宽度是字符，不是像素。而且跟字体有关。宽度1=一个字的1/256.
//            }
//            /**
//             * 设置批注comment
//             */
//            if (titleField != null && StringUtils.isNotEmpty(titleField.getTitleComment())) {
//                anchor.setCol1(cell.getColumnIndex());
//                anchor.setCol2(cell.getColumnIndex() + 1);
//                anchor.setRow1(row.getRowNum());
//                anchor.setRow2(row.getRowNum() + 3);
//
//                // Create the comment and set the text+author
//                Comment comment = drawing.createCellComment(anchor);
//                RichTextString str = factory.createRichTextString(titleField.getTitleComment());
//                comment.setString(str);
//                comment.setAuthor("System");
//                // Assign the comment to the cell
//                cell.setCellComment(comment);
//            }
//        }
        /**
         * 填充数据。
         * sheet对象会被更新。所以没有返回。
         */
        fillData(sheet, data, fieldMap, fieldKeys, rowIndex);
        /**
         * 把文件先写入本地。这个过程，如果数据量巨大，会分批写入。一定程度避免内存溢出。
         */
        File file = null;
        try {
            file = writeFile(workbook, fileName, fileDir);
        } catch (IOException e) {
            logger.error("excel写入磁盘失败！", e);
            throw new ProcessFailedException("excel写入磁盘失败！" + e.getMessage());
        }
        return file;
    }

    /**
     * 创建Excel的标题行。
     * 这个标题行是不带Gaea框架备注的表达式的。供用户使用的。
     * @param sheet
     * @param row
     * @param fieldMap
     * @param drawing
     * @param anchor
     * @param factory
     */
    private void createTitleRow(SXSSFSheet sheet, SXSSFRow row, Map<String, Field> fieldMap, Drawing drawing, ClientAnchor anchor, CreationHelper factory) {
        String[] fieldKeys = fieldMap.keySet().toArray(new String[]{});
        for (int j = 0; j < fieldKeys.length; j++) {
            String fieldKey = fieldKeys[j];
            /**
             * 用结果集data的map的key，去查找title map的key。这两者应该是一致的。
             * 参考CommonViewQueryServiceImpl.query，是经过schemaDataService.transformViewData处理过的。
             */
            Field titleField = fieldMap.get(fieldKey);
            String colTitle = titleField == null ? "" : titleField.getTitle();
            SXSSFCell cell = row.createCell(j);
            cell.setCellValue(colTitle);
            // 设定列宽
            if (NumberUtils.isNumber(titleField.getWidth())) {
                sheet.setColumnWidth(j, Integer.parseInt(titleField.getWidth()) * 256); // 根据API，这里设的宽度是字符，不是像素。而且跟字体有关。宽度1=一个字的1/256.
            }
            /**
             * 设置批注comment
             */
            if (titleField != null && StringUtils.isNotEmpty(titleField.getTitleComment())) {
                anchor.setCol1(cell.getColumnIndex());
                anchor.setCol2(cell.getColumnIndex() + 1);
                anchor.setRow1(row.getRowNum());
                anchor.setRow2(row.getRowNum() + 3);

                // Create the comment and set the text+author
                Comment comment = drawing.createCellComment(anchor);
                RichTextString str = factory.createRichTextString(titleField.getTitleComment());
                comment.setString(str);
                comment.setAuthor("System");
                // Assign the comment to the cell
                cell.setCellComment(comment);
            }
        }
    }

    /**
     * 初始化第一行。创建gaea Excel模板定义说明的行。
     *
     * @param row
     * @param fieldMap
     * @param drawing
     * @param anchor
     * @param factory
     */
    private void createGaeaRow(ExcelSheet excelSheet, SXSSFRow row, Map<String, Field> fieldMap, Drawing drawing, ClientAnchor anchor, CreationHelper factory) {
        // gaea Excel模板定义说明的行不可见。避免给用户困扰。
        row.setZeroHeight(true);

        String[] fieldKeys = fieldMap.keySet().toArray(new String[]{});
        for (int j = 0; j < fieldKeys.length; j++) {
            String fieldKey = fieldKeys[j];
            /**
             * 用结果集data的map的key，去查找title map的key。这两者应该是一致的。
             * 参考CommonViewQueryServiceImpl.query，是经过schemaDataService.transformViewData处理过的。
             */
            Field field = fieldMap.get(fieldKey);

            // ------------------------> 好吧，这段写入单元格内容，其实不是重点
            String colTitle = field == null ? "" : field.getTitle();
            SXSSFCell cell = row.createCell(j);
            cell.setCellValue(colTitle);
            // 非重点 END <------------------------

            // 创建gaea Excel模板定义说明的单元格
            createGaeaCell(excelSheet, row, cell, field, drawing, anchor, factory);
        }
    }

    /**
     * 创建gaea框架的Excel模板单元格/行。以便可以直接用回导出文件再导入。
     *
     * @param row
     * @param cell
     * @param field
     * @param drawing
     * @param anchor
     * @param factory
     */
    private void createGaeaCell(ExcelSheet excelSheet, SXSSFRow row, SXSSFCell cell, Field field, Drawing drawing, ClientAnchor anchor, CreationHelper factory) {
        // 没有ExcelSheet对象就略过
        if (excelSheet == null) {
            return;
        }
        int cellNo = cell.getColumnIndex();
        StringBuilder excelComment = new StringBuilder(); // 某个单元格的备注（这里就是gaea框架的Excel模板定义）
        try {
            // 第一个单元格，需要有sheet定义
            if (cellNo == 0) {
                String sheetDef = ExpressParser.createSheetDefine(excelSheet);
                excelComment.append(sheetDef);
            }
            // 转换生成field定义
            if (field != null) {
                String fieldDef = ExpressParser.createFieldDefine(field);
                excelComment.append(fieldDef);
            }
            // 设定comment位置
            anchor.setCol1(cell.getColumnIndex());
            anchor.setCol2(cell.getColumnIndex() + 5); // 设定comment的宽度
            anchor.setRow1(row.getRowNum());
            anchor.setRow2(row.getRowNum() + 15); // 设定comment的高度

            // Create the comment and set the text+author
            Comment comment = drawing.createCellComment(anchor);
            RichTextString str = factory.createRichTextString(excelComment.toString());
            comment.setString(str);
            comment.setAuthor("System");
            // Assign the comment to the cell
            cell.setCellComment(comment);
        } catch (JsonProcessingException e) {
            logger.error("生成gaea excel模板定义行出错！Json转换失败！", e);
        }
    }

    /**
     * 把workbook写入本地文件。
     *
     * @param workbook
     * @param fileName 文件名。不需要后缀。
     * @param fileDir  要写入的目录。如果结尾没有目录分隔符，自动加上。
     * @return
     * @throws IOException
     */
    private File writeFile(SXSSFWorkbook workbook, String fileName, String fileDir) throws IOException {
        File file = null;
//        String nowTime = DateFormatUtils.format(new Date(), "yyyyMMdd_HHmmss");
        String nowTime = new DateTime().toString("yyyyMMdd_HHmmss");
        // 如果目录的结束没有目录分隔符，加上。
        if (!"/".equals(fileDir.substring(fileDir.length() - 1)) || !"\\".equals(fileDir.substring(fileDir.length() - 1))) {
            fileDir += File.separator;
        }
        if (StringUtils.isEmpty(fileName)) {
            fileName = nowTime + ".xlsx";
        } else {
            if (".xls".equalsIgnoreCase(fileName.substring(fileName.length() - 4)) || ".xlsx".equalsIgnoreCase(fileName.substring(fileName.length() - 5))) {
                logger.warn("XML配置的excel模板的文件名，不需要带文件后缀。fileName='{}'", fileName);
                // 截掉无关的文件后缀
                fileName = fileName.substring(0, fileName.lastIndexOf("."));
            }
            fileName = fileName + "_" + nowTime + ".xlsx";
        }
//            String fullFilePath = "d:\\temp\\excel-export\\" + fileName;
        String fullFilePath = fileDir + fileName;
        FileOutputStream out = new FileOutputStream(fullFilePath);
        file = new File(fullFilePath);
        workbook.write(out);
        out.close();
        return file;
    }

    /**
     * 填充数据到sheet对象中。不返回。
     *  @param sheet
     * @param data
     * @param fieldMap
     * @param fieldKeys
     * @param rowIndex     可以为空。前面已经添加的去到第几行了。
     */
    private void fillData(SXSSFSheet sheet, List<? extends Map> data, Map<String, Field> fieldMap, String[] fieldKeys, Integer rowIndex) throws ValidationFailedException {
        rowIndex = rowIndex==null?0: rowIndex+1;
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        if (ArrayUtils.isEmpty(fieldKeys) && MapUtils.isNotEmpty(fieldMap)) {
            fieldKeys = fieldMap.keySet().toArray(new String[]{});
        }
        // 遍历每一行数据
        for (int i = 0; i < data.size(); i++) {
            Map<String, Object> rowData = data.get(i);
            SXSSFRow row = sheet.createRow(rowIndex+i);
//            String[] fieldKeys = fieldMap.keySet().toArray(new String[]{});
            /**
             * 遍历每一列
             */
            for (int j = 0; j < fieldKeys.length; j++) {
                String fieldKey = fieldKeys[j];
//                    for (String key : rowData.keySet()) {
                Object mapValue = null;
                /**
                 * 假设rowData是LinkedCaseInsensitiveMap，则Key应该是大小写不敏感，用全大写去get也没问题。
                 * 万一传入的map被处理成HashMap之类的，应该默认Key是全大写，也是可以的。
                 * 附：Spring namedParameterJdbcTemplate.queryForList返回的是LinkedCaseInsensitiveMap，这样对于key（数据库列）是大小写不敏感的。如果变成HashMap就会大小写敏感了。
                 */
                mapValue = rowData.get(fieldKey.toUpperCase());
                Field fieldDef = fieldMap.get(fieldKey); // XML 的field定义
                // 创建单元格
                SXSSFCell cell = row.createCell(j);
                Object value = "";
                /**
                 * if 是DataSet的一个结果对象DataItem
                 * 按DataSet方式处理
                 * else
                 */
                if (mapValue instanceof Item) {
                    Item m = (Item) mapValue;
//                        if(titleField.getCellValueTransferBy()==Field.TRANSFER_BY_DS_VALUE) {
                    // 默认值
                    // 避免把null对象转成null字符串
                    value = m.getValue() == null ? "" : m.getValue();
                    if (Field.TRANSFER_BY_DS_TEXT.equalsIgnoreCase(fieldDef.getCellValueTransferBy())) {
                        // 避免把null对象转成null字符串
                        // 如果getText为空（可能DataSet没有对应的转换），则取value的值。
                        value = m.getText() == null ? value : m.getText();
                    }
                } else {
                    // 避免把null对象转成null字符串
                    value = mapValue == null ? "" : mapValue;
                }
                // 设置默认值为字符
                // 如果XML定义的有其他类型，再作转换、覆盖！
                GaeaPoiUtils.setCellValue(cell, value, fieldDef);
            }
        }
    }

    /**
     * 根据Excel Template定义，如果字段有定义其他类型的，例如数字类，则作转换。再把值回写cell。并且把cell的类型设置为对应的类型。
     * 这样，比如对于一些数值类的，在excel里面才可以直接用公式计算。不用再做转换。
     *
     * @param cell
     * @param fieldDef
     * @param value
     */
//    private void parseCellData(SXSSFCell cell, Field fieldDef, String value) {
//        if (Field.DATA_TYPE_NUMBER.equalsIgnoreCase(fieldDef.getDataType()) || Field.DATA_TYPE_DOUBLE.equalsIgnoreCase(fieldDef.getDataType())) {
//            try {
//                cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//                cell.setCellValue(Double.parseDouble(value));
//            } catch (NumberFormatException ex) {
//                // 如果无法转换，可能是带字符还是什么了。直接设为String吧
//                cell.setCellType(XSSFCell.CELL_TYPE_STRING);
//                cell.setCellValue(value);
//            }
//        } else if (Field.DATA_TYPE_INTEGER.equalsIgnoreCase(fieldDef.getDataType())) {
//            try {
//                cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//                cell.setCellValue(Integer.parseInt(value));
//            } catch (NumberFormatException ex) {
//                // 如果无法转换，可能是带字符还是什么了。直接设为String吧
//                cell.setCellType(XSSFCell.CELL_TYPE_STRING);
//                cell.setCellValue(value);
//            }
//        }
//    }
}
