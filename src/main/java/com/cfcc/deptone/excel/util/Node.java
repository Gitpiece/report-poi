package com.cfcc.deptone.excel.util;

/**
 * 
 * @deprecated not use anymore. 
 * @author wanghuanyu
 * @param <T>
 */
@Deprecated
public interface Node<T> {
	public abstract Node<T> getRChild();
	public abstract Node<T> getLChild();
	public abstract void sever();
	public abstract Node<T> getParent();
	public abstract void updateSize();
	public abstract int getSize();
	public abstract void updateHeight();
	public abstract int getHeight();
	public abstract boolean isRChild();
	public abstract boolean isLChild();
	public abstract boolean isLeaf();
	public abstract boolean hasRChild();
	public abstract boolean hasLChild();
	public abstract boolean hasParent();
	public abstract T getData();
	public void setData(T obj) ;
}
