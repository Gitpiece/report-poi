package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;

/**
 * sheet 构建接口
 * 
 * @author WangHuanyu
 */
public interface SheetBuilder {

	/**
	 * 写入数据
	 * @throws POIException
	 */
	void write() throws POIException;

	/**
	 * 初始化写入步骤
	 */
	void initStep();

}
