package test;

import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

/**
 * Created by hOward on 2015/6/13.
 */
public class FormateTest {

    @BeforeClass
    public void before(){
        System.out.println("BeforeClass");
    }

    @Test
    public void test() {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat();
        Date date = new Date();
        System.out.println(Locale.Category.FORMAT);
        System.out.println(Locale.getDefault(Locale.Category.FORMAT));
        System.out.println(simpleDateFormat.format(date));
        simpleDateFormat = new SimpleDateFormat("yyyyyyy-MMM-dddd");
        System.out.println(simpleDateFormat.format(date));

    }
}
