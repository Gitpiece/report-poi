package com.cfcc.deptone.excel.gen;

import com.cfcc.deptone.excel.gen.inner.GeneralExcelBuilder;

public class ExcelBuilderFactory {
	
	private ExcelBuilderFactory(){
		//single
	}
	public static ExcelBuilder getBuilder() {
		return new GeneralExcelBuilder();
	}
}
