package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import org.apache.poi.ss.usermodel.Cell;

import com.cfcc.deptone.excel.model.IConst;
import com.cfcc.deptone.excel.model.ISheet;

public class POIConst extends POICellObject implements IConst {

	private String type;
	private String coordinate;

	protected POIConst(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
		this(poiSheet, cell, arr,true);
	}

	protected POIConst(ISheet poiSheet, Cell cell, String[] arr,boolean emptyV) throws POIException {
		super(poiSheet, cell,emptyV);

		this.coordinate = arr[1];
		this.type = arr[2];
		if(arr.length == 3){
			this.propertyName = "";
		}else{
			this.propertyName = arr[3];
		}
	}

	public String getType() {
		return this.type;
	}

	public String getCoordinate() {
		return this.coordinate;
	}
}
