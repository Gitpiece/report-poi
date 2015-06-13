package com.cfcc.deptone.excel.model;

/**
 * 单元格类型
 * 
 * @author WangHuanyu
 */
public enum CellType {

	HEADER(1, "header"), // 表头
	FOOTER(2, "footer"), // 表底
	NORMAL(3, "normal"), // 普通报表
	SUBTOTAL(31,"subtotal"),//普通报表小计
	SUM(4, "sum"), // 求合
	CONST(5, "const"), // 常量
	CROSSTAB(6, "crosstab"), // 交叉
	CROSSSUM(62, "sum"), // 求和元素
	PICTURE(7, "picture"); // 图片

	private int type;
	private String name;

	CellType(int type, String name) {
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
