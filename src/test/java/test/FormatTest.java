package test;

import com.cfcc.deptone.excel.util.SimpleDecimalFormat;
import org.apache.poi.ss.usermodel.ExcelStyleDateFormatter;
import org.testng.annotations.Test;

import java.math.BigDecimal;
import java.text.*;
import java.util.Date;
import java.util.Locale;

/**
 * FormatTest
 * Created by root on 15-6-14.
 */
public class FormatTest {

//    @Test
    public void Simpledateformattest() {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        System.out.println(012);
        System.out.println(new Date());
        System.out.println(sdf.format(new Date()));
    }

//    @Test
/*    public void testMessageFormat() {
        int fileCount = 1273;
        String diskName = "MyDisk";
        Object[] testArgs = {new Long(fileCount), diskName};

        IMessageFormat form = new IMessageFormat("The disk \"{1}\" contains {0} file(s).");

        System.out.println(form.format(testArgs));

        form = new IMessageFormat("The disk \"{1}\" contains {0}.");
        double[] filelimits = {0,1,2};
        String[] filepart = {"no files","one file","{0,number} files"};
        ChoiceFormat fileform = new ChoiceFormat(filelimits, filepart);
        form.setFormatByArgumentIndex(0, fileform);

        fileCount = 1273;
        diskName = "MyDisk";
        testArgs = new Object[]{new Long(fileCount), diskName};

        System.out.println(form.format(testArgs));
    }
*/
//    @Test
    public void testDecimalFormat(){
        DecimalFormat numberFormat = new DecimalFormat(".00%");
        System.out.println(numberFormat.format(-55.1));
        System.out.println(numberFormat.format(-55.02));
        System.out.println(numberFormat.format(-55.02345));
        numberFormat = new DecimalFormat();
        System.out.println(numberFormat.format(55.02));
    }

    //@Test
    public void testSimpleDecimalFormat(){
        SimpleDecimalFormat simpleDecimalFormat = new SimpleDecimalFormat("#,###.00##");
        BigDecimal bigDecimal = new BigDecimal("12345644584.154");
        System.out.println(simpleDecimalFormat.format(bigDecimal.toString()));

        System.out.println(NumberFormat.getPercentInstance().format(12));
        System.out.println(NumberFormat.getCurrencyInstance().format(12));
        System.out.println(NumberFormat.getNumberInstance().format(113545122.1256));
        System.out.println(NumberFormat.getIntegerInstance().format(113545122.1256));

        //////
        DecimalFormat numberFormat = (DecimalFormat)NumberFormat.getNumberInstance();
        numberFormat.setDecimalSeparatorAlwaysShown(false);
        StringBuffer stringBuffer = new StringBuffer("格式化数字：");
        FieldPosition fieldPosition = new FieldPosition(1);
        fieldPosition.setBeginIndex(1);
        fieldPosition.setEndIndex(4);
        System.out.println(numberFormat.format(12345.,stringBuffer,fieldPosition));
        System.out.println(stringBuffer);
        Double aDouble = new Double(1235.);
        System.out.println(1645.);
    }
    public void testExcelStyleDateFormatter(){
        ExcelStyleDateFormatter excelStyleDateFormatter = new ExcelStyleDateFormatter();
    }

    public static void main(String[] arg){
        String s="(x*y)/z";
//        s.split("[\+\-\\*/()]");
    }
}
