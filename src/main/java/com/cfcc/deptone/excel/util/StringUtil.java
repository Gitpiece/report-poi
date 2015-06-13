package com.cfcc.deptone.excel.util;


import java.math.BigDecimal;
import java.text.MessageFormat;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/**
 * this is a utility for string
 * @author why
 * @version 0.01, 2011/09/27
 *
 */
public class StringUtil {
	private static final Log LOGGER = LogFactory.getLog(StringUtil.class);
	private StringUtil(){
		// do nothing
	}

	public static boolean isEmpty(String arg0){
		return arg0 == null || "".equals(arg0.trim()) ||"null".equalsIgnoreCase(arg0.trim())?true:false;
	}
	
	public static String trim(String arg0){
		return isEmpty(arg0)?"":arg0.trim();
	}
	
	public static Short toShort(String arg0){
		return toShort(arg0,null);
	}
	
	public static Short toShort(String arg0,Short dValue){
		if(isEmpty(arg0)){
			return dValue;
		}else{
			return Short.valueOf(arg0);
		}
	}
	
	public static Integer toInteger(String arg0){
		return toInteger(arg0,null);
	}
	
	public static Integer toInteger(String arg0,Integer dValue){
		if(isEmpty(arg0)){
			return dValue;
		}else{
			return Integer.valueOf(arg0);
		}
	}
	
	public static BigDecimal toBigDecimal(String arg0){
		return toBigDecimal(arg0,null);
	}
	
	public static BigDecimal toBigDecimal(String arg0,BigDecimal dValue){
		if(isEmpty(arg0)){
			return dValue;
		}else{
			return new BigDecimal(arg0);
		}
	}
	
	/**
	 * is digital string,at least 1 length
	 * @param arg0 match String
	 * @param arg1 max digit ,must bigger than 1.
	 * 				default is Integer.MAX_VALUE.
	 * @return
	 */
	public static boolean isDigital(String arg0,Integer ... arg1){
		String max = Integer.MAX_VALUE+"";
		if(arg1.length != 0){
			max = arg1[0].toString();
		}
		
		StringBuilder reg = new StringBuilder("^([0-9]{1,").append(max).append("}+)");
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 是否是报文编号
	 * demo:abcd.001.001.98
	 * @param arg0 number string
	 * @return
	 */
	public static boolean isMsgNumStr(String arg0){
		StringBuilder reg = new StringBuilder();
		reg.append("^([a-z]{4})")
			.append("+(\\.([0-9]{3}))")
			.append("+(\\.([0-9]{3}))")
			.append("+(\\.[0-9]{2})");
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 是否是报文版本
	 * demo:v1.2.2,V1.2.2
	 * @param arg0 number string
	 * @return
	 */
	public static boolean isVersionStr(String arg0){
		StringBuilder reg = new StringBuilder();
		reg.append("^([V,v])")
			.append("+(([0-9]{1}))")
			.append("+(\\.([0-9]{1}))")
			.append("+(\\.[0-9]{1})");
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 校验是否是合法字符，合法字符包括：数字，大小写字母，"."。
	 * @param arg0 stirng(length > 0)
	 * @return
	 */
	public static boolean isLegalStr(String arg0){
		StringBuilder reg = new StringBuilder("^[a-zA-Z0-9.]{1,}");
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 校验是否是WORD文件，文件名以.doc，.docx结尾。
	 * @param arg0 WORD文件名
	 * @return
	 */
	public static boolean isDoc(String arg0){
		StringBuilder reg = new StringBuilder("^(.{0,})+(\\.)+([dD]{1})+([oO]{1})+([cC]{1})+([xX]{0,1})$");//
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 校验是否是rar文件，文件名以.rar结尾。
	 * @param arg0 rar文件名
	 * @return
	 */
	public static boolean isRar(String arg0){
		StringBuilder reg = new StringBuilder("^(.{0,})+(\\.)+([rR]{1})+([aR]{1})+([rR]{1})$");//
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 校验为空
	 * @param arg0 
	 * @return
	 */
	public static boolean isSpace(String arg0){
		StringBuilder reg = new StringBuilder("^([\\s]{0,})$");//
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 校验手机号码
	 * @param arg0 
	 * @return
	 */
	public static boolean isMobilePhone(String arg0){
		StringBuilder reg = new StringBuilder("^(\\+86){0,1}+([\\d]{11})$");//
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 校验电话
	 * @param arg0 
	 * @return
	 */
	public static boolean isTelephone(String arg0){
		StringBuilder reg = new StringBuilder("^((\\d){3,4}-){0,1}+([\\d]{7,8})+(-(\\d){3,4}){0,1}$");//
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 校验电话
	 * @param arg0 
	 * @return
	 */
	public static boolean isContactPhone(String arg0){
		StringBuilder reg = new StringBuilder("(^(\\+86){0,1}+([\\d]{11}))|(^((\\d){3,4}-){0,1}+([\\d]{7,8}))$");//
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * 校验邮箱地址
	 * @param arg0 
	 * @return
	 */
	public static boolean isMail(String arg0){
		String reg="^([\\w-])+(\\.([\\w-])+)*([\\w-])+@([\\w-])+(\\.([\\w-])+)+([\\w-])+$";
		return regularMatch(arg0, reg.toString());
	}
	
	/**
	 * @param arg0
	 * @param reg
	 * @return
	 */
	private static boolean regularMatch(String arg0, String reg) {
		Pattern p = Pattern.compile(reg.toString());
		Matcher m = p.matcher(arg0);
		return m.matches();
	}
	
	//***********************************************************
	//* 格式化
	//***********************************************************
	public static String format(String value, Object[] args) {
		String iString;
		try {
            // ensure the arguments are not null so pre 1.2 VM's don't barf
            Object nonNullArgs[] = args;
            for (int i=0; i<args.length; i++) {
                if (args[i] == null) {
                    if (nonNullArgs==args) {
                    	nonNullArgs=(Object[])args.clone();
                    }
                    nonNullArgs[i] = "null";
                }
            }

            iString = MessageFormat.format(value, nonNullArgs);
        } catch (IllegalArgumentException e) {
        	LOGGER.error(e.getMessage(),e);
            StringBuilder buf = new StringBuilder();
            buf.append(value);
            for (int i = 0; i < args.length; i++) {
                buf.append(" arg[" + i + "]=" + args[i]);
            }
            iString = buf.toString();
            LOGGER.error(iString,e);
        }
		return iString;
	}
}
