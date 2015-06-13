package test;

import com.cfcc.deptone.excel.gen.inner.ExcelBuilderInfo;
import junit.framework.Assert;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by wanghuanyu on 2015/5/19.
 */
public class ExcelBuilderInfoTest {
    @Test
    public void test(){
        List<ExcelBuilderInfo> list = new ArrayList<ExcelBuilderInfo>();
        list.add(new ExcelBuilderInfo(1));
        list.add(new ExcelBuilderInfo(1));
        list.add(new ExcelBuilderInfo(1));
        System.out.println(list.size());
        Assert.assertTrue(list.contains(new ExcelBuilderInfo(1)));
    }
}
