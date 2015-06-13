package com.cfcc.deptone.excel.model;


/**
 * excel sheet的枚举类型
 *
 * @author WangHuanyu
 */
public enum SheetType {

    NORMAL_SHEET(1, "normla"), //普通报表
    CROSSTAB_SHEET(2, "cross tab"), //交叉报表
    CUST_1(3, "header,footer,cost and pic");//模板只包含表头、表尾、常量、图片

    private int type;
    private String name;

    SheetType(int type, String name) {
        this.type = type;
        this.name = name;
    }

    public int getType() {
        return type;
    }

    public String getName() {
        return name;
    }
}
