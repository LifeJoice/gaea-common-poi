package org.gaea.poi;

import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.domain.GaeaPoiFieldDefine;
import org.gaea.poi.domain.GaeaPoiResultGroup;
import org.gaea.poi.domain.GaeaPoiResultSet;
import org.gaea.poi.util.ExpressParser;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeanWrapper;
import org.springframework.beans.PropertyAccessorFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.text.MessageFormat;
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

    public <T> GaeaPoiResultGroup<T> getData(InputStream fileIS, Class<T> beanClass) throws ValidationFailedException {
        GaeaPoiResultGroup<T> resultGroup = new GaeaPoiResultGroup<T>();
        GaeaPoiResultSet<T> resultSet = new GaeaPoiResultSet<T>();
        T bean = BeanUtils.instantiate(beanClass);
        BeanWrapper wrapper = PropertyAccessorFactory.forBeanPropertyAccess(bean);
        wrapper.setAutoGrowNestedPaths(true);
        Map<String,String> excelRowValue = new HashMap<String, String>();
        Map<Integer,GaeaPoiFieldDefine> columnDefMap = new HashMap<Integer, GaeaPoiFieldDefine>();
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
                    if(cell.getRowIndex()==0) {
                        String myExcelRemark = cell.getCellComment().getString().getString();
                        GaeaPoiFieldDefine fieldDef = expressParser.parseField(myExcelRemark);
                        columnDefMap.put(cell.getColumnIndex(),fieldDef);
                    }
                    String value = getCellStringValue(cell);
                    String key = columnDefMap.get(cell.getColumnIndex()).getName();
                    excelRowValue.put(key,value);
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

    private String getCellStringValue(Cell cell) {
        String value = "";
        if(cell.getCellType()==Cell.CELL_TYPE_STRING){
            value = cell.getStringCellValue();
        }else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
            if(HSSFDateUtil.isCellDateFormatted(cell)){
                value = cell.getDateCellValue().toString();
                System.out.println(DateFormatUtils.format(cell.getDateCellValue().getTime(),"yyyy-MM-dd HH:mm:ss"));
                return value;
            }
            value = String.valueOf(cell.getNumericCellValue());
        }
        return value;
    }
}
