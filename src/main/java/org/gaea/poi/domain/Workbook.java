package org.gaea.poi.domain;

import java.util.List;

/**
 * Created by iverson on 2016-6-6 16:33:39.
 */
public class Workbook {
    private String id;
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
