package com.cfcc.deptone.excel.data;

import org.apache.commons.beanutils.DynaBean;
import org.apache.commons.beanutils.LazyDynaBean;

import java.io.File;
import java.math.BigDecimal;
import java.util.*;

public class TestData {

    public static final String resourcePath = System.getProperty("user.dir")
            + File.separator + "src\\test\\resource" + File.separator;
    public static final String reportPath = System.getProperty("user.dir")
            + File.separator + "target" + File.separator;

    public static Map<String, Object> getMetadata() {
        final Map<String, Object> _map = new HashMap<String, Object>();
        _map.put("table", "制表：中心主管");
        _map.put("top", "1900/1/7");
        _map.put("cachet", "国库公章：无");
        _map.put("strcode", "全国");
        _map.put("treasuryname", "中国人民银行");
        _map.put("director", "业务主管：无");
        _map.put("titledate", "2011年06月30日<日报>");
        _map.put("date", "2012-2-23");
        _map.put("date1", new Date());
        _map.put("check", "复核：小飞侠");
        _map.put("rptime", "日");
        _map.put("HIDDEN_DATA", "null,0");
        _map.put("newdate", "2011年06月30日<日报>");
        _map.put("scope", "全辖");
        _map.put("bdglevel", "按空格选择可多选");
        _map.put("amtunit", "元");
        _map.put("tatol", new BigDecimal("234564.12"));
        _map.put("Boolean", new Boolean(false));
        _map.put("Float", new Float(12.2));

        return _map;
    }

    public static List<TempObj> getNormalList() {
        List<TempObj> list = new ArrayList<TempObj>();
        TempObj _tmp = new TempObj();
        _tmp.setSbtcode("101");
        _tmp.setName("税收收入");
        _tmp.setNavelamt(0);
        _tmp.setNavelyearamt(0);
        _tmp.setLocalamt(new BigDecimal(237808580649999999999999999.07));
        _tmp.setLocalyearamt(111625891111.23);
        _tmp.setSUM_amt(0.07);
        _tmp.setSUM_yearamt(111625891111.23);
        list.add(_tmp);

        _tmp = new TempObj();
        _tmp.setSbtcode("101");
        _tmp.setName("增值税");
        _tmp.setNavelamt(1);
        _tmp.setNavelyearamt(0);
        _tmp.setLocalamt(new BigDecimal(2378085806.07));
        _tmp.setLocalyearamt(11162589111.23);
        _tmp.setSUM_amt(0.67);
        _tmp.setSUM_yearamt(11162589111.23);
        list.add(_tmp);

        _tmp = new TempObj();
        _tmp.setSbtcode("1010101");
        _tmp.setName("国内增值税");
        _tmp.setNavelamt(3);
        _tmp.setNavelyearamt(0);
        _tmp.setLocalamt(new BigDecimal(237808580.07));
        _tmp.setLocalyearamt(1116258911.23);
        _tmp.setSUM_amt(1.0);
        _tmp.setSUM_yearamt(1116258911.23);
        list.add(_tmp);

        _tmp = new TempObj();
        _tmp.setSbtcode("1010101");
        _tmp.setName("国内增值税");
        _tmp.setNavelamt(3);
        _tmp.setNavelyearamt(0);
        _tmp.setLocalamt(new BigDecimal(Float.MAX_VALUE));
        _tmp.setLocalyearamt(1116258911.23);
        _tmp.setSUM_amt(1.0);
        _tmp.setSUM_yearamt(1116258911.23);
        list.add(_tmp);
        return list;
    }

    /**
     * 交叉报表测试数据
     *
     * @return
     */
    public static List<TempObj> getCrossList() {
        List<TempObj> list = new ArrayList<TempObj>();
        TempObj _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本期执行数");
        _tmp.setTitle3("金额");
        _tmp.setName("中央总库");
        _tmp.setSbtcode("101");
        _tmp.setEndamt(23780858064.07);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本期执行数");
        _tmp.setTitle3("同比±(%)");
        _tmp.setName("中央总库");
        _tmp.setSbtcode("101");
        _tmp.setEndamt(50.0);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本年累计数");
        _tmp.setTitle3("金额");
        _tmp.setName("中央总库");
        _tmp.setSbtcode("101");
        _tmp.setEndamt(237808580.07);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本年累计数");
        _tmp.setTitle3("同比±(%)");
        _tmp.setName("中央总库");
        _tmp.setSbtcode("101");
        _tmp.setEndamt(60.0);
        _tmp.setRate(20.00);
        list.add(_tmp);

        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本期执行数");
        _tmp.setTitle3("金额");
        _tmp.setName("中央总库");
        _tmp.setSbtcode("10101");
        _tmp.setEndamt(101010858064.07);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本期执行数");
        _tmp.setTitle3("同比±(%)");
        _tmp.setName("中央总库");
        _tmp.setSbtcode("10101");
        _tmp.setEndamt(52.0);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本年累计数");
        _tmp.setTitle3("金额");
        _tmp.setName("中央总库");
        _tmp.setSbtcode("10101");
        _tmp.setEndamt(10101.9);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本年累计数");
        _tmp.setTitle3("同比±(%)");
        _tmp.setName("中央总库");
        _tmp.setSbtcode("10101");
        _tmp.setEndamt(62.0);
        _tmp.setRate(20.00);
        list.add(_tmp);

        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本期执行数");
        _tmp.setTitle3("金额");
        _tmp.setName("北京");
        _tmp.setSbtcode("101");
        _tmp.setEndamt(13780858064.07);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本期执行数");
        _tmp.setTitle3("同比±(%)");
        _tmp.setName("北京");
        _tmp.setSbtcode("101");
        _tmp.setEndamt(51.0);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本年累计数");
        _tmp.setTitle3("金额");
        _tmp.setName("北京");
        _tmp.setSbtcode("101");
        _tmp.setEndamt(137808580.07);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本年累计数");
        _tmp.setTitle3("同比±(%)");
        _tmp.setName("北京");
        _tmp.setSbtcode("101");
        _tmp.setEndamt(61.0);
        _tmp.setRate(20.00);
        list.add(_tmp);

        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本期执行数");
        _tmp.setTitle3("金额");
        _tmp.setName("北京");
        _tmp.setSbtcode("10101");
        _tmp.setEndamt(13780064.07);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本期执行数");
        _tmp.setTitle3("同比±(%)");
        _tmp.setName("北京");
        _tmp.setSbtcode("10101");
        _tmp.setEndamt(53.9);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本年累计数");
        _tmp.setTitle3("金额");
        _tmp.setName("北京");
        _tmp.setSbtcode("10101");
        _tmp.setEndamt(18580.07);
        _tmp.setRate(20.00);
        list.add(_tmp);
        _tmp = new TempObj();
        _tmp.setTitle1("一般预算收入");
        _tmp.setTitle2("本年累计数");
        _tmp.setTitle3("同比±(%)");
        _tmp.setName("北京");
        _tmp.setSbtcode("10101");
        _tmp.setEndamt(63.9);
        _tmp.setRate(20.00);
        list.add(_tmp);
        return list;
    }

    public static void main(String[] args) {
        System.out.println(resourcePath);
    }


    public static List<DynaBean> getListFor1row1column() {
        List<DynaBean> list = new ArrayList<DynaBean>();
        list.add(create1row1columnBean("山东高铁","2015-05-19","23.6"));
        list.add(create1row1columnBean("靖远煤电","2015-05-19","11.61"));
        list.add(create1row1columnBean("山东高铁","2015-05-20","23.62"));
        list.add(create1row1columnBean("靖远煤电","2015-05-20","11.78"));
        return list;
    }

    private static DynaBean create1row1columnBean(String colname1, String rowname1, String value) {
        DynaBean myBean = new LazyDynaBean();
        myBean.set("colname1",colname1);
        myBean.set("rowname1",rowname1);
        myBean.set("value",value);
        return myBean;
    }
}
