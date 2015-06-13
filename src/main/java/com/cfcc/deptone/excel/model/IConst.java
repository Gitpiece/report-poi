package com.cfcc.deptone.excel.model;

/**
 * 静态单元格接口
 * 
 * @author WangHuanyu
 */
public interface IConst extends ICellObject {


	String getType();

	/**
	 * 得到坐标常量
	 * @return 坐标
	 */
	String getCoordinate();
}
