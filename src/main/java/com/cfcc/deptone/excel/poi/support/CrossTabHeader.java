package com.cfcc.deptone.excel.poi.support;

import com.cfcc.deptone.excel.model.ICrossTab;

/**
 * 交叉报表行列表头
 * @author wanghuanyu
 *
 */
public class CrossTabHeader {
	/**POI模板对象*/
	private ICrossTab crossTabPOI;
	/**表头内容*/
	private String title;
	private String propertyName;
	
	public CrossTabHeader(ICrossTab crossTabPOI,String title){
		this.crossTabPOI = crossTabPOI;
		this.title = title;
	}
	
	//=================================
	//  	临时文档
	public CrossTabHeader(String propertyName,String title){
		this.propertyName = propertyName;
		this.title = title;
	}
}
