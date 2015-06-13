package com.cfcc.deptone.excel.model;

/**
 * 求合操作单元格接口
 * 
 * @author WangHuanyu
 */
public interface IOpSum extends ICellObject {

	/**
	 * 合的名字，在一个sheet中是唯一的。
	 * 这个名字可以用来求合表达式
	 * @return
	 */
	String getName();

	String getCoordinate();

	/**
	 * 是否含有表达式
	 * @return
	 */
	boolean hasMember();

	/**
	 * 返回定义的表达式
	 * @return
	 */
	String getMembers();
	
	/**
	 * 设置初始坐标
	 * @param oc
	 */
	void setOriginalColumn(int oc);
	
	void getOriginalRow(int or);
}
