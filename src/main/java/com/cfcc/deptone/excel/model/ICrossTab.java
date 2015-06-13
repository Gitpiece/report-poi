package com.cfcc.deptone.excel.model;

/**
 * 交叉报表单元格接口
 * 
 * @author WangHuanyu
 */
public interface ICrossTab extends ICellObject {

	/**列占位符*/
	String CROSSTAB_COLUMN = "column";
	/**行占位符*/
	String CROSSTAB_ROW = "row";
	/**数据占位符*/
	String CROSSTAB_DATA = "data";

	/**
	 * 是列
	 * @return 是否交叉报表列
	 */
	boolean isColumn();
	
	/**
	 * 是行
	 * @return 是否交叉报表行
	 */
	boolean isRow();
	
	/**
	 * 是数据
	 * @return 是否是交叉报表数据
	 */
	boolean isData();
	/**
	 * 行、列返回实际层级，如：1，2，3，列自上而下递增，行从左到右递增。数据返回0。
	 * @return 层级
	 */
	int level();

}
