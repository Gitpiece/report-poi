package com.cfcc.deptone.excel;

import com.cfcc.deptone.excel.util.TreeNode;

public class TreeNodeTest {

	@org.junit.Before
	public void before() {

	}

	@org.junit.Test
	public void test() {
		TreeNode<String> root = new TreeNode<String>();
		root.setData("root");
		root.setLabel("root");
		System.out.println(root.getData());
//		System.out.println(root.getSize());
		
		TreeNode<String> A = new TreeNode<String>(root,"A","A");
		TreeNode<String> B = new TreeNode<String>(root,"B","B");
		TreeNode<String> C = new TreeNode<String>(root,"C","C");
//		System.out.println(root.getSize());
		root.removeChild(A);
//		System.out.println(root.getSize());
		
		org.junit.Assert.assertNotNull(A.getParent());
		
		root.insertChild(A);
//		System.out.println(root.getSize());
		A.sever();
//		System.out.println(root.getSize());
		
	}
	
}
