package com.cfcc.deptone.excel.poi.operation;

/**
 * 合并操作操作
 * 代码示例：<code>merge [v|h] [merge number]</code>
 * 暂时未实现h。
 * 暂时未实现merge number选项如果不填写，合并检查到的所有内容相同的横向或纵向单元格；
 * 如果填写，检查merge number个单元格，如果单元格内容相同，合并单元格，如果不同，不进行合并操作。
 *
 * @author wanghuanyu
 */
public class POIMergeOperation extends POIOperation {

    public static final String OPERATION = "merge";

    /**
     * 合并类型类型，纵向v，或横向h
     */
    private String mergeType;

    /**
     * 合并的行数或列数。
     * 如果不填写
     */

    public String getMergeType() {
        return mergeType;
    }

    public void setMergeType(String mergeType) {
        this.mergeType = mergeType;
    }

}
