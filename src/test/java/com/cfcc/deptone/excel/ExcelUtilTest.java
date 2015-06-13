package com.cfcc.deptone.excel;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.util.POIExcelUtil;

/**
 * Created by wanghuanyu on 2015/6/9.
 */
public class ExcelUtilTest {

    public void mergeExcelTest() throws POIException {
        POIExcelUtil.mergeExcel(new String[]{"D:\\temp\\merge\\source-a.xlsx", "D:\\temp\\merge\\source-b.xlsx", "D:\\temp\\merge\\监控统计报表.xlsx"}, "D:\\temp\\merge\\creatReport-merge.xlsx");
    }

    public void mergeExcelTest2003() throws POIException {
        POIExcelUtil.mergeExcel(new String[]{"D:\\temp\\merge\\source-a.xls", "D:\\temp\\merge\\source-b.xls", "D:\\temp\\merge\\监控统计报表.xls"}, "D:\\temp\\merge\\creatReport-merge.xls");
    }
}
