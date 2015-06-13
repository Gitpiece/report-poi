package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import org.apache.poi.ss.usermodel.Cell;

import com.cfcc.deptone.excel.model.IFooter;
import com.cfcc.deptone.excel.model.ISheet;

public class POIFooter extends POICellObject implements IFooter {

	protected POIFooter(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
		super(poiSheet, cell);
		this.propertyName = arr[1];
	}

}
