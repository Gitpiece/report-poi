package com.cfcc.deptone.excel;

import java.io.File;
import java.io.IOException;
import java.util.Map;

import org.apache.commons.io.FileUtils;

import com.cfcc.deptone.excel.data.TestData;
import com.cfcc.deptone.excel.gen.ExcelBuilderFactory;
import com.cfcc.deptone.excel.gen.ExcelBuilder;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class NormalDoubleinputTest {
	Map<String, Object> metadata;
	@BeforeTest
	public void beforetest() throws IOException{
		metadata = TestData.getMetadata();
		metadata.put("num", "#{normal.SN}");
		metadata.put("sbtcode", "#{normal.sbtcode}");
		metadata.put("name", "#{normal.name}");
		metadata.put("navelamt", "#{normal.navelamt}");
		metadata.put("navelyearamt", "#{normal.navelyearamt}");
		metadata.put("localamt","#{normal.localamt}");
		metadata.put("localyearamt","#{normal.localyearamt}");	
		metadata.put("SUM_amt","#{normal.SUM_amt}");
		metadata.put("SUM_yearamt","#{normal.SUM_yearamt}");
		
		//footer
		metadata.put("cachetp","#{footer.cachet}");
		metadata.put("directorp","#{footer.director}");
		metadata.put("checkp","#{footer.check}");
		metadata.put("tablep","#{footer.table}");
		
		FileUtils.copyFile(new File(TestData.resourcePath + "normal-doubleinput.xls"),new File(TestData.reportPath + "normal-doubleinput-out.xls"));
		FileUtils.copyFile(new File(TestData.resourcePath + "normal-doubleinput.xlsx"),new File(TestData.reportPath + "normal-doubleinput-out.xlsx"));
	}
	
	@Test
	public void testNormal2003() throws Exception {
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		excelBuilder.build(TestData.reportPath+"normal-doubleinput-out.xls", metadata, TestData.getNormalList());
		excelBuilder.build(TestData.reportPath+"normal-doubleinput-out.xls", metadata, TestData.getNormalList());
	}
	@Test
	public void testNormal2007() throws Exception {
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		excelBuilder.build(TestData.reportPath+"normal-doubleinput-out.xlsx", metadata, TestData.getNormalList());
		excelBuilder.build(TestData.reportPath+"normal-doubleinput-out.xlsx", metadata, TestData.getNormalList());
	}
	
}