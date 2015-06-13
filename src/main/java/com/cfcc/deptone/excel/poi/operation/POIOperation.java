package com.cfcc.deptone.excel.poi.operation;

/**
 * Poi扩展操作参数
 *
 * @author wanghuanyu
 */
public abstract class POIOperation {
    /**
     * 操作类型
     */
    private String operation;

    /**
     *
     * @return 操作类型
     */
    public String getOperation() {
        return operation;
    }

    public void setOperation(String operation) {
        this.operation = operation;
    }
}
