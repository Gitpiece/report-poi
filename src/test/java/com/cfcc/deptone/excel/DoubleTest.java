package com.cfcc.deptone.excel;

import java.math.BigDecimal;

import com.cfcc.deptone.excel.util.POIExcelUtil;

public class DoubleTest {

	@org.junit.Test
	public void test(){
		Double d = new Double(125D);
		System.out.println(Float.MAX_VALUE);
		BigDecimal bd = new BigDecimal(2378085806412654123.07);
		System.out.println(POIExcelUtil.formateBigDecimal(new BigDecimal(Float.MAX_VALUE)));
	}
}
