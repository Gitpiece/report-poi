package com.cfcc.deptone.excel.model;

/**
 * 小计
 *
 * @author WangHuanyu
 */
public interface ISubTotal extends ICellObject {
    /**
     * 类型
     *
     * @return 小计类型
     * @see com.cfcc.deptone.excel.util.ExcelConsts
     */
    String getType();

    /**
     * <code>com.cfcc.deptone.excel.util.ExcelConsts.STATIC</code> 类型的值
     *
     * @return
     */
    String getValue();
}
