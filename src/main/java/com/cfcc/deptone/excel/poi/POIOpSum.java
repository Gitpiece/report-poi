package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import org.apache.poi.ss.usermodel.Cell;

import com.cfcc.deptone.excel.model.IOpSum;
import com.cfcc.deptone.excel.model.ISheet;

public class POIOpSum extends POICellObject implements IOpSum {

	private String sumName;
	private String coordinate;
	boolean hasMember = false;
	String members;

	protected POIOpSum(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
		this(poiSheet,cell,arr,true);
	}
	
	protected POIOpSum(ISheet poiSheet, Cell cell, String[] arr,boolean emptyV) throws POIException {
		super(poiSheet, cell,emptyV);

		this.coordinate = arr[1];
		this.sumName = arr[2];
		if (arr.length >= 4) {
			members = arr[3];
			hasMember = true;
		}
	}

	public String getCoordinate() {
		return this.coordinate;
	}

	public String getName() {
		return this.sumName;
	}

	public String getMembers() {
		return this.members;
	}

	public boolean hasMember() {
		return hasMember;
	}

	public void setOriginalColumn(int oc) {
		super.originalColumn = oc;
	}

	public void getOriginalRow(int or) {
		super.originalRow = or;
	}

}
