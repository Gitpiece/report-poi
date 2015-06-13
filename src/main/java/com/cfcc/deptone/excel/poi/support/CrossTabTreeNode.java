package com.cfcc.deptone.excel.poi.support;

import com.cfcc.deptone.excel.model.ICrossTab;
import com.cfcc.deptone.excel.util.TreeNode;

public class CrossTabTreeNode<T> extends TreeNode<T> {
	/**POI模板对象*/
	private ICrossTab crossTab;
	public CrossTabTreeNode(){
		super();
	}
	
	public CrossTabTreeNode(CrossTabTreeNode<T> parent,String label, T e) {
		super(parent,label, e);
	}

	public ICrossTab getCrossTab() {
		return crossTab;
	}

	public void setCrossTab(ICrossTab crossTab) {
		this.crossTab = crossTab;
	}
	
	
}
