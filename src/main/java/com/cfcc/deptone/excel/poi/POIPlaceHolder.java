package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.util.POIExcelUtil;

public class POIPlaceHolder implements IPlaceHolder {
    private String placeHolderStr;

    protected POIPlaceHolder(String placeHolderStr) {
        this.placeHolderStr = placeHolderStr;
    }

    public String toPHString() {
        return this.placeHolderStr;
    }

    public String[] toArray() {
        return POIExcelUtil.toStrArray(placeHolderStr);
    }

}
