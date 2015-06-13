package com.cfcc.deptone.excel.poi.operation;

/**
 * 排序操作
 * 代码示例：<code>order [asc|desc]</code>
 * @author wanghuanyu
 *
 */
public class POIOrderOperation extends POIOperation {
	
	public static final String OPERATION = "order";
	public static final String ORDER_TYPE_ASC = "asc";
	public static final String ORDER_TYPE_DESC = "desc";
	
	/**排序类型，升序asc，或降序desc*/
	private String orderType;

	public String getOrderType() {
		return orderType;
	}

	public void setOrderType(String orderType) {
		this.orderType = orderType;
	}
}
