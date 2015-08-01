package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ISheet;

/**
 * sheet 构建步骤接口
 * 
 * @author WangHuanyu
 */
public interface BuildStep {

	/**
	 * 步骤名称
	 * 
	 * @return
	 */
	String getName();

	/**
	 * 构建sheet步骤执行方法
	 * @param sheet
	 * @throws Exception
	 */
	void build(ISheet sheet) throws POIException;

	/**
	 * 写入row之前操作
	 * @throws Exception
	 */
	void beforeRow() throws POIException;

	/**
	 * 写入row之后操作
	 * @throws Exception
	 */
	void afterRow() throws POIException;
}
