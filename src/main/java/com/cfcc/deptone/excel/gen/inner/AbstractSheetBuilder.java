package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.model.ISheet;

public abstract class AbstractSheetBuilder implements SheetBuilder {

	protected ISheet sheet;

	public AbstractSheetBuilder(ISheet sheet) {
		this.sheet = sheet;
	}

	public abstract void initStep();
}
