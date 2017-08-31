package org.gaea.poi.domain;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by iverson on 2016-6-4 16:21:39.
 */
public class ExcelSheet implements Serializable{
    private String id;// excel定义的id，对应的是XML配置的<workbook>的id
    private List<ExcelBlock> blockList;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public List<ExcelBlock> getBlockList() {
        if(blockList==null){
            blockList = new ArrayList<ExcelBlock>();
        }
        return blockList;
    }

    public void setBlockList(List<ExcelBlock> blockList) {
        this.blockList = blockList;
    }
}
