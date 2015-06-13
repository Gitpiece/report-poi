package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import org.apache.poi.ss.usermodel.Cell;

import com.cfcc.deptone.excel.model.IHeader;
import com.cfcc.deptone.excel.model.ISheet;

public class POIHeader extends POICellObject implements IHeader {

	protected POIHeader(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
		super(poiSheet, cell);
		this.propertyName = arr[1];
	}

}
