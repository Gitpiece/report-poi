import org.testng.annotations.Test;

import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Scanner;

/**
 * We promptly judged antique ivory buckles for the next prize
 * pangram
 * We promptly judged antique ivory buckles for the prize
 * not pangram
 *
 * @author wanghuanyu
 */
public class SolutionTest {

    @Test
    public void test() {
        try {
            System.out.println(URLEncoder.encode("jdbc:db2://10.1.3.31:49500/aiispdb:currentSchema=aiis","utf-8"));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        Scanner in = new Scanner(System.in);
        String input = in.nextLine();

        int length = input.length();
        int i = 0;

        Calendar cal = new GregorianCalendar();


    }
}
