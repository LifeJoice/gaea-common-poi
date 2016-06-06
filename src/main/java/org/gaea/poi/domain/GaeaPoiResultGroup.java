package org.gaea.poi.domain;

import java.util.List;

/**
 * Created by iverson on 2016-6-4 16:47:19.
 */
public class GaeaPoiResultGroup<T> {
    private List<GaeaPoiResultGroup> subGroups;
    private GaeaPoiResultSet<T> result;

    public List<GaeaPoiResultGroup> getSubGroups() {
        return subGroups;
    }

    public void setSubGroups(List<GaeaPoiResultGroup> subGroups) {
        this.subGroups = subGroups;
    }

    public GaeaPoiResultSet<T> getResult() {
        return result;
    }

    public void setResult(GaeaPoiResultSet<T> result) {
        this.result = result;
    }
}
