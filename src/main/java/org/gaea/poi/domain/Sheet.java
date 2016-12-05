package org.gaea.poi.domain;

import java.io.Serializable;
import java.util.List;

/**
 * 实现Serializable是为了可以缓存对象。
 * Created by iverson on 2016-11-2 19:48:40.
 */
public class Sheet implements Serializable{
    private String id;// excel定义的id，对应的是XML配置的<workbook>的id
    private List<Block> blockList;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public List<Block> getBlockList() {
        return blockList;
    }

    public void setBlockList(List<Block> blockList) {
        this.blockList = blockList;
    }
}
