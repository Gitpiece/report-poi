package com.cfcc.deptone.excel.model;

/**
 * 常规单元格接口
 * #{normal.sbtcode.[merge 1]}
 * @author WangHuanyu
 */
public interface INormal extends ICellObject {

	/**合并开关，后面需要紧跟一个整型参数。*/
	String OPTIONAL_MERGE = "merge";
	/* =====================================================
	 * 常量属性， 一般定义通用的填写值，如：自增序号，每写一行自动增加1；
	 * 用法：只需要把Dto中的属性名替换为常量名即可，如：#{normal.SN}，会从一开始计数自增；
	 =====================================================*/
	/** 自增序号属性常量	 */
	String SN = "SN";
	
	/*=====================================================
	 * 
	 * 常规单元格借口参数
	 * 
	 =====================================================*/
	/**
	 * 是否启用合并开关
	 * <p>
	 * 选项增加[merge n]
	 */
	boolean isOpMerge();
	/**
	 * 当 {@link #isOpMerge()} 开关打开时，本参数才生效，合并开关格检测个数。返回数字体现的是合并单元格个数（包括当前单元格）.
	 * 所以程序会检测当前单元格之后N-1个单元格的内容是否与当前单元格内容一致(equals)，如果一致合并单元格，否则忽略。
	 * @return 合并单元格数
	 */
	int opMergeNumber();

	/**
	 * 是否是分组，序结合subtotal使用。
	 * @return
	 */
	boolean hasGroup();
}
