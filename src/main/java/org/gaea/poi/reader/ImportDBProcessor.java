package org.gaea.poi.reader;

import org.apache.commons.lang3.StringUtils;
import org.gaea.db.ibatis.jdbc.SQL;
import org.gaea.exception.ValidationFailedException;
import org.gaea.poi.domain.Block;
import org.gaea.poi.domain.ExcelField;
import org.gaea.poi.domain.Field;
import org.gaea.poi.domain.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import javax.sql.DataSource;
import java.text.MessageFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 数据导入处理类。当需要把Excel直接导入数据库的时候，使用本类。
 * Created by iverson on 2016-6-6 11:40:53.
 */
@Component
public class ImportDBProcessor {
    private final Logger logger = LoggerFactory.getLogger(ImportDBProcessor.class);
    @Autowired(required = false)
    private DataSource dataSource;
    @Autowired(required = false)
    private NamedParameterJdbcTemplate namedParameterJdbcTemplate;

    /**
     * 初始化导入功能
     * 主要初始化数据库连接。因为需要直接从excel读取写入数据库。必须检查系统的数据库配置是否完整。
     */
    @PostConstruct
    protected void init(){
        if(namedParameterJdbcTemplate==null){
            logger.debug("无法获取Spring容器的NamedParameterJdbcTemplate。尝试创建自己的NamedParameterJdbcTemplate。");
            if(dataSource==null){
                logger.warn("通用Excel导入功能初始化。无法获取Spring容器的NamedParameterJdbcTemplate和datasource，直接导入数据库将不可用！");
            }else{
                namedParameterJdbcTemplate = new NamedParameterJdbcTemplate(dataSource);
                logger.debug("无法获取Spring容器的NamedParameterJdbcTemplate。尝试创建自己的 NamedParameterJdbcTemplate 完成。");
            }
        }
    }

    public void executeImport(Block<Map<String,String>> block) throws ValidationFailedException {
        if(block==null || block.getData()==null){
            throw new ValidationFailedException("Excel文档数据（block）为空，无法执行导入！");
        }
        if(StringUtils.isEmpty(block.getTable())){
            throw new ValidationFailedException("Excel通用导入配置不完整，无法执行导入！缺失关键配置项block table.");
        }
        String tableName = block.getTable();
        StringBuilder columns = new StringBuilder();// 插入SQL的列。例如：id,name,address
        StringBuilder values = new StringBuilder();// 插入SQL的值(占位符).例如: :id,:name,:address
        // 遍历XML定义的字段（和Excel定义应该已经合并过了），拼凑插入列和对应的占位符遍历
        for(Field field:block.getFieldMap().values()){
            columns.append(field.getName()).append(",");
            String key = field.getName();
            String paramName = ":"+key;// PreparedStatement和spring的SQL占位符
            values.append(paramName).append(",");
        }
        String strColumns = StringUtils.removeEnd(columns.toString(),",");
        String strValues = StringUtils.removeEnd(values.toString(),",");
        // 构造SQL
        String insertSQL = new SQL().INSERT_INTO(tableName).VALUES(strColumns,strValues).toString();
        logger.debug(MessageFormat.format("insert SQL:\n{0}",insertSQL).toString());
        namedParameterJdbcTemplate.batchUpdate(insertSQL,block.getData().toArray(new HashMap[0]));
    }
}
