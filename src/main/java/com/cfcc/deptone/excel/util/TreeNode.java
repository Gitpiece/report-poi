package com.cfcc.deptone.excel.util;

import com.cfcc.deptone.excel.poi.operation.POIOrderOperation;

import java.util.Arrays;
import java.util.Comparator;

/**
 * 有序树树数据类型。
 * <br /><br />
 * 局限性：同辈的label必须唯一。
 * @author wanghuanyu
 * @param <T>
 */
public class TreeNode<T> {
	/**数据域*/
	protected T data; 
	/**父结点，树根没有父节点*/
	private TreeNode<T> parent; 

	/**子节点数组，有序*/
	protected TreeNode<T>[] childs = null;
	/**初始化容量*/
	protected int initialCapacity = 5;
	/**扩展容量*/
	protected int increaseCapacity = 10;
	/**子孙数（包括结点本身）。*/
	/**size = 0,*/
	/**高度*/
	protected int height = 0;
	/**树的度，树包含子树的个数。度为0的节点为叶子节点*/
	protected int degree = 0;
	
	/**节点标签，在同辈中唯一*/
	protected String label;

	/**
	 * create root
	 */
	public TreeNode() {
		this(null, "/", null);
	}

	/**
	 * 创建节点
	 * @param parent 父
	 * @param label 标签
	 * @param e 数据对象
	 */
	public TreeNode(TreeNode<T> parent, String label, T e) {
		this.setParent(parent);
		if (this.getParent() != null) {
			this.getParent().insertChild(this);
		}

		setLabel(label);
		setData(e);
		/**this.size = 1;*/
		this.childs = (TreeNode<T>[]) new TreeNode[this.initialCapacity];
	}

	/****** Node 接口方法 ******/
	/**
	 * 获取节点数据
	 * @return
	 */
	public T getData() {
		return this.data;
	}

	/**
	 * 设置节点数据
	 * @param obj
	 */
	public void setData(T obj) {
		data = obj;
	}

	/**
	 * 设置标签
	 * @param label
	 */
	public void setLabel(String label) {
		this.label = label;
	}

	public String getLabel() {
		return this.label;
	}

	/****** 辅助方法,判断当前结点位置情况 ******/
	/**
	 * 判断是否有父亲
	 * @return
	 */
	public boolean hasParent() {
		return this.getParent() != null;
	}

	/**
	 * 返回树根
	 * @return
	 */
	public TreeNode<T> getRoot() {
		TreeNode<T> root = this;
		while (true) {
			TreeNode<T> tmp = root.getParent();
			if (tmp == null) {
				break;
			}
		}

		return root;
	}

	/**
	 * 判断是否为叶子结点
	 * @return
	 */
	public boolean isLeaf() {
		return hasChild();
	}

	/**
	 * 是否有子孙
	 * @return
	 */
	public boolean hasChild() {
		return this.degree>0;
	}

	/**
	 * 返回树的度
	 * @return
	 */
	public int getDegree() {
		return this.degree;
	}

	/**
	 * 增加存储容量
	 * @param minCapacity
	 */
	public void ensureCapacity(int minCapacity) {
		int oldCapacity = childs.length;
		if (minCapacity > oldCapacity) {
			int newCapacity = (oldCapacity * 3) / 2 + 1;
			if (newCapacity < minCapacity) {
				newCapacity = minCapacity;
			}
			// minCapacity is usually close to size, so this is a win:
			childs = ArraysUtil.copyOf(childs, newCapacity);
		}
	}

	/****** 与height 相关的方法 ******/
	/**
	 *  取结点的高度,即以该结点为根的树的高度
	 *  TODO 待完善
	 * @return
	 */
	public int getHeight() {
		return height;
	}

	/**
	 * 返回节点的层，跟节点是0层。
	 * @return
	 */
	public int getLevel() {
		int level = 0;

		while (true) {
			if (this.getParent() == null) {
				break;
			}
			level++;
		}

		return level;
	}

	/**
	 * 更新当前结点及其祖先的高度
	 */
	protected void updateHeight() {
		int newH = 0;// 新高度初始化为0,高度等于左右子树高度加1 中的大者
		for (int i = 0; i < this.degree; i++) {
			newH = Math.max(newH, 1 + childs[i].getHeight());
		}
		// 高度没有发生变化则直接返回
		if (newH == height) {
			return;
		}

		// 否则更新高度
		height = newH;
		if (hasParent()) {
			getParent().updateHeight(); // 递归更新祖先的高度
		}
	}

	/****** 与size 相关的方法 ******/
	/**
	 * 取以该结点为根的树的结点数
	 * @return
	 */

	/**
	 * 更新当前结点及其祖先的子孙数
	 */

	/****** 与parent 相关的方法 ******/
	/**
	 * 取父结点
	 * @return
	 */
	public TreeNode<T> getParent() {
		return this.parent;
	}
	protected void setParent(TreeNode<T> p) {
		parent = p;
	}

	/**
	 * 断开与父亲的关系
	 */
	public void sever() {
		if (!hasParent()) {
			return;
		}
		this.getParent().removeChild(this);
		this.setParent(null);
		
		//parent.updateHeight(); // 更新父结点及其祖先高度
		//parent.updateSize(); // 更新父结点及其祖先规模
	}

	//===================与Child 相关的方法 
	/**
	 * 移除子节点
	 * @param tn 移除节点
	 * @return 移除节点
	 */
	public TreeNode<T> removeChild(TreeNode<T> tn) {
		TreeNode<T> tmp = null;
		//查找需要移除的节点
		int i;
		for (i = 0; i < this.getAllChild().length; i++) {
			TreeNode<T> treeNode = childs[i];
			if (treeNode.getLabel().equals(tn.getLabel())) {
				tmp = treeNode;
				break;
			}
		}
		//移动节点兄弟
		for (i++; i < this.getAllChild().length; i++) {
			childs[i - 1] = childs[i];
		}
		childs[this.getAllChild().length] = null;
		this.degree--;

		return tmp;
	}

	/**
	 * 插入节点
	 * @param child
	 */
	public void insertChild(TreeNode<T> child) {
		if (this.degree == childs.length) {
			ensureCapacity(childs.length + this.increaseCapacity);
		}
		childs[this.degree++] = child;
	}

	/**
	 * 返回所有子节点
	 * @return
	 */
	public TreeNode<T>[] getAllChild() {
		return ArraysUtil.copyOf(childs, this.degree);
	}

	/**
	 * 返回子节点
	 * @param label
	 * @return
	 */
	public TreeNode<T> getChild(String label) {
		for (int i = 0; i < this.degree; i++) {
			TreeNode<T> treeNode = childs[i];
			if (treeNode.getLabel().equals(label)) {
				return treeNode;
			}
		}
		return null;
	}

	/**
	 * 返回子节点
	 * @param index 子节点索引
	 * @return
	 */
	public TreeNode<T> getChild(int index) {
		//如果超过子树范围，返回空
		if (index > this.degree) {
			return null;
		}
		return childs[index];
	}

	/**
	 * 是否包含节点，以label为判断依据。
	 * @param label
	 */
	public boolean containNode(String label) {
		for (int i = 0; i < this.degree; i++) {
			TreeNode<T> treeNode = childs[i];
			if (treeNode.getLabel().equals(label)) {
				return true;
			}
		}
		return false;
	}
	
	/**
	 * 排序子节点
	 * @param ordertype
	 */
	@SuppressWarnings("unchecked")
	public void orderChild(String ordertype){
		TreeNode<T>[] array = this.getAllChild();
		if(POIOrderOperation.ORDER_TYPE_ASC.equals(ordertype)){
			Arrays.sort(array,new AscOrder());
		} else {
			Arrays.sort(array,new DescOrder());
		}
		
		this.childs = array;
	}
	
	class AscOrder implements Comparator{
		
		public int compare(Object o1, Object o2) {
			return ((TreeNode)o1).getLabel().compareTo(((TreeNode)o2).getLabel());
		}
		
	}
	class DescOrder extends AscOrder{

		public int compare(Object o1, Object o2) {
			return -super.compare(o1, o2);
		}
		
	}
}
