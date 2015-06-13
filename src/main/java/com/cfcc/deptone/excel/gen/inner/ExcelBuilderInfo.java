package com.cfcc.deptone.excel.gen.inner;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.Map;

/**
 * 报表构建信息，包括excel sheet页及所标信息，构建所需要的数据。
 */
public class ExcelBuilderInfo {
    /**
     * sheet索引，从0开始。
     */
    private int index;
    /**
     * excel sheet对象
     */
    private Sheet sheet;
    /**
     * 报表数据
     */
    private Map metaDataMap;
    /**
     * 报表数据
     */
    private List rptDataList;

    public ExcelBuilderInfo(int index) {
        this(index, null, null, null);
    }

    public ExcelBuilderInfo(int index, Sheet sheet, Map metaDataMap, List rptDataList) {
        this.index = index;
        this.sheet = sheet;
        this.metaDataMap = metaDataMap;
        this.rptDataList = rptDataList;
    }

    /**
     * 是否包含数据
     *
     * @return
     */
    public boolean hasData() {
        return this.metaDataMap != null;
    }

    @Override
    public int hashCode() {
        return this.index;
    }

    @Override
    public boolean equals(Object obj) {
        ExcelBuilderInfo to = (ExcelBuilderInfo) obj;
        return to.hashCode() == this.hashCode();
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public Map getMetaDataMap() {
        return metaDataMap;
    }

    public void setMetaDataMap(Map metaDataMap) {
        this.metaDataMap = metaDataMap;
    }

    public List getRptDataList() {
        return rptDataList;
    }

    public void setRptDataList(List rptDataList) {
        this.rptDataList = rptDataList;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }
}