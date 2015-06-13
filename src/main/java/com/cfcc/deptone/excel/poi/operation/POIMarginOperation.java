package com.cfcc.deptone.excel.poi.operation;

/**
 * 排序操作
 * 代码示例：<code>group</code>
 * @author wanghuanyu
 *
 */
public class POIMarginOperation extends POIOperation {
	
	public static final String OPERATION = "margin";

	/**排序类型，升序center,left,right*/
	private String margintype;
	/**
	 * 只有margintype为left,right才会有此属性
	 */
	private Integer number ;

	public String getMargintype() {
		return margintype;
	}

	public Integer getNumber() {
		return number;
	}

	public void setMargintype(String margintype) {
		this.margintype = margintype;
	}

	public void setNumber(Integer number) {
		this.number = number;
	}
}
