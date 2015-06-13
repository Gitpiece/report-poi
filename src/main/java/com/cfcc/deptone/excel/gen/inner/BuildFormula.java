package com.cfcc.deptone.excel.gen.inner;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IFormula;
import com.cfcc.deptone.excel.model.ISheet;

/**
 * 重新设置模板中的公式
 * @author WangHuanyu
 */
public class BuildFormula implements BuildStep {
	public String getName() {
		return "BuildFormula";
	}

	public void build(ISheet sheet) throws POIException {
		List<IFormula> allOpSum = sheet.getAllFormula();
		for (IFormula iFormula : allOpSum) {
			Cell cell  = sheet.getCell(iFormula.getOriginalRow(), iFormula.getOriginalColumn());
			cell.setCellFormula(iFormula.getCellFormula());
		}

	}

	public void afterRow() throws POIException {
		//do nothking
	}

	public void beforeRow() throws POIException {
		//do nothking
	}

}
