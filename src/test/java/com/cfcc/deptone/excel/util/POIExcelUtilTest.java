package com.cfcc.deptone.excel.util;

import com.cfcc.deptone.excel.data.TestData;
import com.cfcc.deptone.excel.gen.POIException;
import org.testng.annotations.Test;

import java.io.File;

/**
 * Created by wanghuanyu on 2015/6/23.
 */
public class POIExcelUtilTest {

    @Test
    public void getAllSheetNameTest(){
        File file = new File(TestData.resourcePath + "normal.xls");
        try {
            String[] sheetnames = POIExcelUtil.getAllSheetName(file.getAbsolutePath());
            System.out.println("sheet names:");
            for (String sheetname: sheetnames){
                System.out.format("%15s",sheetname);
            }
        } catch (POIException e) {
            e.printStackTrace();

        }
    }
}
