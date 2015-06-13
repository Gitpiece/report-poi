package com.cfcc.deptone.excel.poi.support;

import com.cfcc.deptone.excel.model.ICrossTab;

/**
 * 交叉报表数据
 * @author wanghuanyu
 *
 */
public class CrossTabData {
	/**POI模板对象*/
	private ICrossTab crossTabDataPOI;
	/**原始数据对象*/
	private Object originalObj;
	Point point;
	/**行表头，按树形结构保存，如：root>sub，便于查找节点坐标。*/
	private String rowString;
	/**列表头，按树形结构保存，如：root>sub，便于查找节点坐标。*/
	private String columnString;
	public ICrossTab getCrossTabDataPOI() {
		return crossTabDataPOI;
	}
	public void setCrossTabDataPOI(ICrossTab crossTabDataPOI) {
		this.crossTabDataPOI = crossTabDataPOI;
	}
	public Object getOriginalObj() {
		return originalObj;
	}
	public void setOriginalObj(Object originalObj) {
		this.originalObj = originalObj;
	}
	public String getRowString() {
		return rowString;
	}
	public void setRowString(String rowString) {
		this.rowString = rowString;
	}
	public String getColumnString() {
		return columnString;
	}
	public void setColumnString(String columnString) {
		this.columnString = columnString;
	}
	
	
	public void setPoint(int x, int y) {
		this.point = new Point(x,y);
	}
	public Point getPoint() {
		return point;
	}
	
	/**
	 * 坐标类
	 * @author wanghuanyu
	 */
	public class Point {
		int x, y;

		protected Point(int x, int y) {
			this.x = x;
			this.y = y;
		}

		public int getX() {
			return x;
		}

		public int getY() {
			return y;
		}
	}

}
