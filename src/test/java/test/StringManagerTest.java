package test;

import com.cfcc.deptone.excel.util.StringManager;
import org.testng.annotations.Test;

/**
 * StringManagerTest
 * Created by wanghuanyu on 2015/6/15.
 */
public class StringManagerTest {

    @Test
    public void test(){
        StringManager sm = StringManager.getManager("com.cfcc.deptone.excel.gen.inner");
        System.out.println(sm.getString("poi.test"));
        System.out.println(sm.getString("poi.filenotfound"));
    }
}
