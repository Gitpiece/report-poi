package com.cfcc.deptone.excel.poi.operation;

import com.cfcc.deptone.excel.poi.POICrossSum;

/**
 * 传播操作。
 * 比如交叉报表 {@link POICrossSum} 的求和，需要根据写入的数据进行横向或纵向的自动传播。
 * 代码示例：<code>spread [all|odd|even]</code>，all可以省略。
 * @author wanghuanyu
 *
 */
public class SpreadOperation extends POIOperation {
	
	public static final String OPERATION = "spread";

	//奇数传播类型
	public static final String SPREAD_TYPE_ODD = "odd";
	//偶数传播类型
	public static final String SPREAD_TYPE_EVEN = "even";
	//全部传播
	public static final String SPREAD_TYPE_ALL ="all";
	
	/**
	 * 传播类型。
	 */
	private String spreadType;

	public String getSpreadType() {
		return spreadType;
	}


	public void setSpreadType(String spreadType) {
		this.spreadType = spreadType;
	}
	
}
