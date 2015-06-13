package com.cfcc.deptone.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;

import com.cfcc.deptone.excel.data.TestData;
import com.cfcc.deptone.excel.gen.ExcelBuilderFactory;
import com.cfcc.deptone.excel.gen.ExcelBuilder;
import org.testng.annotations.Test;

public class Cust1Test {

	@Test
	public void tesetCustRpt2003() throws Exception{
		FileUtils.copyFile(new File(TestData.resourcePath + "cust1.xls"),new File(TestData.reportPath + "cust1-out.xls"));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		excelBuilder.build(TestData.reportPath+"cust1-out.xls", getMetadata(), getCrossList());
	}
	@Test
	public void tesetCustRpt2007() throws Exception{
		FileUtils.copyFile(new File(TestData.resourcePath + "cust1.xlsx"),new File(TestData.reportPath + "cust1-out.xlsx"));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		excelBuilder.build(TestData.reportPath+"cust1-out.xlsx", getMetadata(), getCrossList());
	}

	private Map<String, Object> getMetadata() {
		Map<String,Object> map = new HashMap<String,Object>();
		map.put("pic", TestData.resourcePath+"fish.gif");
		map.put("name", "小春");
		map.put("gender", "男");
		map.put("resume", "1.人见人爱\n2.马桶\n3.人的大脑\n");
	    return map;
    }

	private List getCrossList() {
	    return new ArrayList();
    }
}
