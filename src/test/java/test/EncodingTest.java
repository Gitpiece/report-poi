package test;

import org.testng.annotations.Test;

import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.net.URLEncoder;

/**
 * Created by admin on 2016/1/8.
 */
public class EncodingTest {

    @Test
    public void test() throws UnsupportedEncodingException {
        System.out.println(URLDecoder.decode("\\xB7\\XB4","utf-8"));
    }
}
