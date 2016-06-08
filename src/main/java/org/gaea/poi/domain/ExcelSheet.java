package org.gaea.poi.domain;

/**
 * Created by iverson on 2016-6-4 16:21:39.
 */
public class ExcelSheet {
    private String id;// excel定义的id，对应的是XML配置的<workbook>的id
    private ExcelBlock block;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public ExcelBlock getBlock() {
        return block;
    }

    public void setBlock(ExcelBlock block) {
        this.block = block;
    }
}
