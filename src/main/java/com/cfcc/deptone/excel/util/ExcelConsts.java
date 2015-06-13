/**
 *
 */
package com.cfcc.deptone.excel.util;


/**
 * 2012-3-23
 * @author WangHuanyu
 */
public class ExcelConsts {
    private ExcelConsts() {
        //single
    }

    /**
     * excel数值有效位数，超过此位数的数值，末尾会抹零。
     */
    public final static int NUMERICAL_VALID_DIGIT = 15;

    /* 单元格格式，设置单元格格式->数字中的分类*/
    public static final String CELL_FORMATE_GENERAL = "General";//常规
    public static final String CELL_FORMATE_TEXT = "@";//文本

    public static final String VERTICAL = "v";
    public static final String HORIZONTAL = "h";
    public static final String FIX = "f";

    /**
     * 常量cell类型，propertyname既是单元格需要写入的值。
     */
    public static final String STATIC = "static";//propertyname就是单元格需要写入的值。
    /**
     * 参数cell类型，通过propertyname在元集合中得到value。
     */
    public static final String PARAM = "param";//需要用propertyname在sheet的元集合中得到value

    //*****************************************************************
    // subtotal operation type
    //*****************************************************************
    /**
     * 小计计算类型
     */
    public static final String CALC = "calc";

    //*****************************************************************
    // margin operation type
    //*****************************************************************
    /**
     * margin left
     */
    public static final String MARGIN_CENTER = "center";
    /**
     * margin left
     */
    public static final String MARGIN_LEFT = "left";
     /**
     * margin right
     */
    public static final String MARGIN_RIGHT= "right";

    //兼容报表组件参数
    /**
     * <p>
     * 为了兼容报表组件，与 com.cfcc.deptone.excel.util.ExcelConsts.REPORT_TEMPLATE_SHEET_NAME 值一样。<p>
     * </p>
     * metadata中的参数，指定模板sheet页名称。示例：
     * <p><blockquote><pre>
     *	Map metadata = new HashMap();
     *	metadata.put(ExcelConsts.REPORT_TEMPLATE_SHEET_NAME, "sheetname");
     *	</pre></blockquote><p>
     */
    public static final String REPORT_TEMPLATE_SHEET_NAME = "REPORT_TEMPLATE_NAME_KEY";
}
