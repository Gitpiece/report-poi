package com.cfcc.deptone.excel.model.support;

/**
 * 类似Spring的初始化bean
 *
 * @author wanghuanyu
 */
public interface InitBean {
    /**
     * 初始化后操作
     */
    void afterInit();
}
