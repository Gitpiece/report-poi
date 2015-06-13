package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import org.apache.poi.ss.usermodel.Cell;

import com.cfcc.deptone.excel.model.IPicture;
import com.cfcc.deptone.excel.model.ISheet;

public class POIPicture extends POICellObject implements  IPicture {

	protected POIPicture(ISheet sheet, Cell cell,String[] arr) throws POIException {
	    super(sheet, cell);
	    
		this.propertyName = arr[1];
    }


}
