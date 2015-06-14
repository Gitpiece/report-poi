package com.cfcc.deptone.excel.util;

import java.text.FieldPosition;
import java.text.ParsePosition;

/**
 * Created by wanghuanyu on 2015/6/14.
 */
public class SimpleDecimalFormat {//extends Format{


    // Constants for characters used in programmatic (unlocalized) patterns.
    private static final char       PATTERN_ZERO_DIGIT         = '0';
    private static final char       PATTERN_GROUPING_SEPARATOR = ',';
    private static final char       PATTERN_DECIMAL_SEPARATOR  = '.';
    private static final char       PATTERN_DIGIT              = '#';
    private static final char       PATTERN_PERCENT            = '%';


    private int      left_digit_groupsize   = 0;
    private int      right_zero_digit       = 0;
    private int      right_digit            = 0;
    private boolean  percent                = false;

    private String pattern;
    public SimpleDecimalFormat(String pattern){
        this.pattern = pattern;

        applyPattern(this.pattern);
    }

    private void applyPattern(String pattern){

        int digitleftcount=0,percentcount=0,digitleftgroupsize=0;
        boolean left = true,right = false,ingroup = false;

        int pattenlength = pattern.length();
        for (int pos = 0; pos < pattenlength ; pos++) {
            char ch = pattern.charAt(pos);
            switch (ch) {
                case PATTERN_ZERO_DIGIT:
                    if(left){
                        digitleftcount++;
                    }
                    if(right && right_digit>0){
                        throw new IllegalArgumentException("Unexpected '0' in pattern \"" +
                                pattern + '"');
                    }
                    right_zero_digit++;
                    break;
                case PATTERN_GROUPING_SEPARATOR:
                    if(digitleftgroupsize > left_digit_groupsize){
                        left_digit_groupsize = digitleftgroupsize;
                    }
                    ingroup = true;
                    digitleftgroupsize = 0;
                    break;
                case PATTERN_DECIMAL_SEPARATOR:
                    if(digitleftcount > 0){
                        throw new IllegalArgumentException("Unexpected '0' in pattern \"" +
                                pattern + '"');
                    }
                    if(digitleftgroupsize > left_digit_groupsize){
                        left_digit_groupsize = digitleftgroupsize;
                    }
                    left = false;
                    right = true;
                    break;
                case PATTERN_DIGIT:
                    if(ingroup){
                        digitleftgroupsize++;
                    }
                    if(right){
                        right_digit ++;
                    }
                    break;
                case PATTERN_PERCENT:
                    percentcount++;
                    break;
                default:
                    throw new IllegalArgumentException("Unexpected '"+ch+"' in pattern \"" +
                            pattern + '"');
            }
        }

        if(percentcount >1){
            throw new IllegalArgumentException("Multiple percent in pattern \"" +
                    pattern + '"');
        }else if(percentcount == 1){
            percent = true;
        }


    }


    public StringBuilder format(Object obj){
        if(obj instanceof String){
            return this.formatString((String)obj, new StringBuilder(), new FieldPosition(0));
        }

        return this.format(obj,new StringBuilder(),new FieldPosition(0));

    }

    private StringBuilder formatString(String obj, StringBuilder stringBuilder, FieldPosition fieldPosition) {
        int beginIndex = fieldPosition.getBeginIndex();
        int endIndex = fieldPosition.getEndIndex();
        int dotindex = obj.lastIndexOf(PATTERN_DECIMAL_SEPARATOR);
        StringBuilder lefttemp = new StringBuilder(obj.substring(0, dotindex));
        StringBuilder left = new StringBuilder();
        String righttemp = obj.substring(dotindex+1,obj.length());
        StringBuilder right = new StringBuilder();

        if(percent){
            int off = righttemp.length()<2?righttemp.length():2;
            lefttemp.append(righttemp.substring(1,off));
            righttemp =righttemp.substring(off+1,righttemp.length());
        }
        //left
        int notfillcount = lefttemp.length()%left_digit_groupsize;
        int looptime = (lefttemp.length())/left_digit_groupsize;
        left.append(lefttemp.substring(lefttemp.length(),notfillcount));
        for (int i = 0; i < looptime; i++) {
            left.append(PATTERN_GROUPING_SEPARATOR).append(lefttemp.substring(notfillcount + (i + 1) * left_digit_groupsize, (i + 1) * left_digit_groupsize+notfillcount));
        }

        //right
        right.append(righttemp);
        if(right.length() < right_zero_digit){
            int supple = right_zero_digit-right.length();
            for (int i = 0; i < supple; i++) {
                right.append(PATTERN_ZERO_DIGIT);
            }
        }

        stringBuilder.append(right).append(PATTERN_DECIMAL_SEPARATOR).append(right);
        if(percent)stringBuilder.append(PATTERN_PERCENT);
        return stringBuilder;
    }

    public StringBuilder format(Object obj, StringBuilder toAppendTo, FieldPosition pos) {
        return null;
    }

    private StringBuilder subformat(StringBuilder result,
                                   boolean isNegative, boolean isInteger,
                                   int maxIntDigits, int minIntDigits,
                                   int maxFraDigits, int minFraDigits) {
        return null;
    }

    public Object parseObject(String source, ParsePosition pos) {
        return null;
    }
}
