package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.model.ISubTotal;
import com.cfcc.deptone.excel.util.ExcelConsts;
import org.apache.poi.ss.usermodel.Cell;

public class POISubTotal extends POICellObject implements ISubTotal {

	private String type;
	private String value="";

	protected POISubTotal(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
		this(poiSheet,cell,arr,true);
	}

	protected POISubTotal(ISheet poiSheet, Cell cell, String[] arr, boolean emptyV) throws POIException {
		super(poiSheet, cell,emptyV);

		this.type = arr[1];
		if (ExcelConsts.STATIC.equals(this.type) && arr.length >2) {
			value = arr[2];
		}
	}

	public String getType() {
		return type;
	}

	public String getValue() {
		return value;
	}
}
