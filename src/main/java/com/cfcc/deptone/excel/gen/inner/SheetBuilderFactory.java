package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.model.SheetType;
import com.cfcc.deptone.excel.util.StringManager;

/**
 * 报表构建工场，生产不同类型报表的构建类。
 * date 2012-4-1
 * @author WangHuanyu
 */
public class SheetBuilderFactory {
	protected static StringManager sm = StringManager.getManager("com.cfcc.deptone.excel.gen.inner");
	private SheetBuilderFactory() {
		//single
	}

	public static SheetBuilder getBuilder(ISheet sheet) throws POIException {
		if (SheetType.NORMAL_SHEET.getType() == sheet.getSheetType().getType() ) {
			return new SheetBuilderNormal(sheet);
		} else if (SheetType.CROSSTAB_SHEET.getType() == sheet.getSheetType().getType() ) {
			return new SheetBuilderCrossTab(sheet);
		} else if (SheetType.CUST_1.getType() == sheet.getSheetType().getType()) {
			return new SheetBuilderCust1(sheet);
		}

		throw new POIException(sm.getString("poi.template.sheet.buildernotfound", sheet.getSheet().getSheetName()));
	}

}
