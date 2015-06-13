package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ICrossTab;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.model.ISheet;
import org.apache.poi.ss.usermodel.Cell;

import java.util.List;

public class POICrossTab extends POICellObject implements ICrossTab {

	protected String crossType;
	int level = 0;

	protected POICrossTab(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
		super(poiSheet, cell);
		
		List<IPlaceHolder> list = this.getAllPlaceHolder();
		crossType = list.get(0).toArray()[1];
	}

	public String getPropertyName() {
		List<IPlaceHolder> list = this.getAllPlaceHolder();
		return list.get(0).toArray()[2];
	}

	public boolean isColumn() {
		return ICrossTab.CROSSTAB_COLUMN.equals(crossType);
	}

	public boolean isRow() {
		return ICrossTab.CROSSTAB_ROW.equals(crossType);
	}
	
	public boolean isData() {
		return ICrossTab.CROSSTAB_DATA.equals(crossType);
	}

	public int level() {
		return level;
	}
}
