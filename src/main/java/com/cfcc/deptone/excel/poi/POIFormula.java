package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import org.apache.poi.ss.usermodel.Cell;

import com.cfcc.deptone.excel.model.IFormula;
import com.cfcc.deptone.excel.model.ISheet;

public class POIFormula extends POICellObject implements IFormula {
	private String formula;
	protected POIFormula(ISheet sheet, Cell cell) throws POIException {
	    super(sheet, cell);
	    this.formula = cell.getCellFormula(); 
    }
	
	public String getCellFormula(){
		return this.formula;
	}

}
