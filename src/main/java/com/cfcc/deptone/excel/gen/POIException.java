package com.cfcc.deptone.excel.gen;

/**
 * poi exception
 * @author wanghuanyu
 *
 */
public class POIException extends Exception {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public POIException(String string, IndexOutOfBoundsException e) {
		super(string, e);
	}

	public POIException(Exception e) {
		super(e);
	}
	
	public POIException(String msg) {
		super(msg);
	}

}
