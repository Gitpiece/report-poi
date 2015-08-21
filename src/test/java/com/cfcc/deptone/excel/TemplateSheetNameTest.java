package com.cfcc.deptone.excel;

import com.cfcc.deptone.excel.data.TestData;
import com.cfcc.deptone.excel.gen.ExcelBuilderFactory;
import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.util.ExcelConsts;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * 指定模板sheet name 测试
 * Created by wanghuanyu on 2015/5/20.
 */
public class TemplateSheetNameTest {

    @Test
    public void testSheetName1row1column() throws IOException, POIException {
        FileUtils.copyFile(new File(TestData.resourcePath + "creatReport-poi.xls"), new File(TestData.reportPath + "creatReport-poi-out.xls"));

        Map map1 = TestData.getMetadata();
        Map map = TestData.getMetadata();
        map.put("title","测试单行单列报表");
        map.put(ExcelConsts.REPORT_TEMPLATE_SHEET_NAME,"oneRowOneColumn");
        ExcelBuilderFactory.getBuilder().build(TestData.reportPath + "creatReport-poi-out.xls", new Map[]{map1,map}, new List[]{TestData.getListFor1row1column(),TestData.getListFor1row1column()});
    }

    @Test
    public void testSheetName() throws IOException, POIException {
        FileUtils.copyFile(new File(TestData.resourcePath + "normal.xls"), new File(TestData.reportPath + "normal-sheetname-out.xls"));

        Map map1 = TestData.getMetadata();
        map1.put(ExcelConsts.REPORT_TEMPLATE_SHEET_NAME,"sheet3");
        Map map = TestData.getMetadata();
        map.put("title","测试sheetname");
        map.put(ExcelConsts.REPORT_TEMPLATE_SHEET_NAME,"sheet2");
        ExcelBuilderFactory.getBuilder().build(TestData.reportPath + "normal-sheetname-out.xls", new Map[]{map1,map}, new List[]{TestData.getNormalList(),TestData.getNormalList()});
    }
}
