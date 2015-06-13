package com.cfcc.deptone.excel.model;

/**
 * 交叉求和单元格接口
 * 
 * @author WangHuanyu
 */
public interface ICrossSum extends ICellObject {

	/**
	 * @deprecated replaced by {@value #SPREAD}
	 */
	@Deprecated
	String EXPAND = "expand";
	String SPREAD = "spread";
	
	/**
	 * 得到坐标，表示求和单元格是横向求和还是纵向求和。
	 * <p>
	 * {@link com.cfcc.deptone.excel.util.ExcelConsts#HORIZONTAL} or {@link com.cfcc.deptone.excel.util.ExcelConsts#VERTICAL}
	 * @return 坐标
	 */
	String getCoordinate();
	
	/**
	 * 是否展开
	 * @deprecated replaced by {@link #isSpread()}
	 * @return 是否展开
	 */
	@Deprecated
	boolean isExpand();

	/**
	 * 是否有传播性
	 * @return 是否有传播性
	 */
	boolean isSpread();
	
	/**
	 * 传播类型
	 * @return 传播类型
	 */
	String getSpreadType();
}
