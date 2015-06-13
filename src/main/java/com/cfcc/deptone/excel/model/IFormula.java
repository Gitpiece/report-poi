package com.cfcc.deptone.excel.model;

/**
 * excel公式类型单元格
 * 
 * @author WangHuanyu
 */
public interface IFormula extends ICellObject {
	
	/**
	 * 得到表达式
	 * @return 公式
	 */
	String getCellFormula();

}
