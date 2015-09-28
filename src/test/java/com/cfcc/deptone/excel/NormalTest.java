package com.cfcc.deptone.excel;

import com.cfcc.deptone.excel.data.TestData;
import com.cfcc.deptone.excel.gen.ExcelBuilder;
import com.cfcc.deptone.excel.gen.ExcelBuilderFactory;
import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class NormalTest {

	@Test
	public void testNormal2003() throws Exception {
		POIExcelUtil.copyFile(new File(TestData.resourcePath + "normal.xls"), new File(TestData.reportPath + "normal-out.xls"));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		excelBuilder.build( TestData.reportPath+"normal-out.xls", TestData.getMetadata(), TestData.getNormalList());
	}

	@Test
	public void testNormal2003writetwice() throws Exception {
		String reportfile = TestData.reportPath + "normal-twice-out.xls";
		POIExcelUtil.copyFile(new File(TestData.resourcePath + "normal-twice.xls"), new File(reportfile));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		Map<String,Object> map = TestData.getMetadata();
		map.put("sn","#{normal.SN}");
		map.put("sbtcode","#{normal.sbtcode}");
		map.put("name","#{normal.name}");
		map.put("localamt","#{normal.localamt}");
		map.put("localyearamt","#{normal.localyearamt}");
		excelBuilder.build(reportfile, map, TestData.getNormalList());
		excelBuilder.build(reportfile, map, TestData.getNormalList());
	}

	@Test
	public void testNormal20034ChineseName() throws Exception {
		String template = "normal中文.xls";
		String report = "normal-out中文.xls";
		POIExcelUtil.copyFile(new File(TestData.resourcePath + template), new File(TestData.reportPath + report));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		excelBuilder.build( TestData.reportPath+report, TestData.getMetadata(), TestData.getNormalList());
	}

	/**
	 * 当map中传递了sheetname并且并且在模板中不存在的情况
	 * @throws Exception
	 */
	@Test
	public void testNormal2003_2() throws POIException{
		String outputFileName = "normal-out2.xls";
		POIExcelUtil.copyFile(new File(TestData.resourcePath + "normal.xls"), new File(TestData.reportPath + outputFileName));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		Map map = TestData.getMetadata();
		//map.put(ExcelConsts.REPORT_TEMPLATE_SHEET_NAME,"notexistsheet");
		try {
			excelBuilder.build( TestData.reportPath+outputFileName, map, TestData.getNormalList());
		} catch(IllegalArgumentException e){
			Assert.assertTrue(true);
		} catch (POIException e) {
			throw e;
		}
	}

	/**
	 * 测试普通报表。共三个sheet页，其中第1,3 sheet页的数据为Null。
	 * @throws Exception
	 */
	@Test
	public void testNormal2003_multsheet() throws Exception {
		POIExcelUtil.copyFile(new File(TestData.resourcePath + "normal.xls"), new File(TestData.reportPath + "normal-outm.xls"));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		List[] date = new ArrayList[3];
		Map[] metadate = new HashMap[3];

		date[1] = TestData.getNormalList();
		metadate[1] = TestData.getMetadata();
		excelBuilder.build(TestData.reportPath + "normal-outm.xls", metadate, date);
	}

	@Test
	public void testNormal2007() throws Exception {
		POIExcelUtil.copyFile(new File(TestData.resourcePath + "normal.xlsx"), new File(TestData.reportPath + "normal-out.xlsx"));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		excelBuilder.build(TestData.reportPath+"normal-out.xlsx", TestData.getMetadata(), TestData.getNormalList());
	}

	@Test
	public void testNormal2003_subtotal() throws Exception {
		POIExcelUtil.copyFile(new File(TestData.resourcePath + "normal-subtotal.xls"), new File(TestData.reportPath + "normal-subtotal-out.xls"));
		ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
		excelBuilder.build( TestData.reportPath+"normal-subtotal-out.xls", TestData.getMetadata(), TestData.getNormalList());
	}


}