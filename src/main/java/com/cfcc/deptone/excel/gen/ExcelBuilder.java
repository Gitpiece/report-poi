package com.cfcc.deptone.excel.gen;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Excel报表构建接口
 * 
 * 2012-3-23
 * @author why
 */
public interface ExcelBuilder {

	/**
	 * 构建Excel报文
	 * 
	 * @param reportPath 生成报表文件路径
	 * @param metaDataMap 元数据集合，以key-value形式保存，应用于表头、表尾、常量、图片等类型的单元格
	 * @param rptDataList 报表数据集合
	 * @return 0:成功，-1失败
	 * @throws POIException
	 */
	int build(String reportPath, Map<String, Object> metaDataMap[], List<?> rptDataList[]) throws POIException;
	int build(String reportPath, Map<String, Object> metaDataMap, List<?> rptDataList) throws POIException;

	int build(File reportPath, Map<String, Object> metaDataMap[], List<?> rptDataList[]) throws POIException;
	int build(File reportPath, Map<String, Object> metaDataMap, List<?> rptDataList) throws POIException;
	
}
