package com.cfcc.deptone.excel.gen.inner;

import org.apache.poi.ss.util.CellRangeAddress;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IConst;
import com.cfcc.deptone.excel.model.ISheet;

/**
 * 删除常量合并信息步骤
 * 
 * @author WangHuanyu
 */
public class BuildRemoveConstMerger implements BuildStep {
	public String getName() {
		return "BuildRemoveOpSumMerger";
	}

	public void build(ISheet sheet) throws POIException {

		for (IConst iConst : sheet.getAllConst()) {

			for (int i = sheet.getNumMergedRegions(); i > 0; i--) {
				CellRangeAddress rangeAddress = sheet.getMergedRegion(i - 1);

				if (iConst.getOriginalRow() >= rangeAddress.getFirstRow() && iConst.getOriginalRow() <= rangeAddress.getLastRow()
				    && iConst.getOriginalColumn() >= rangeAddress.getFirstColumn() && iConst.getOriginalColumn() <= rangeAddress.getLastColumn()) {

					iConst.setRangeAddress(rangeAddress);
					sheet.removeMergedRegion(i - 1);
				}
			}
		}
	}

	public void afterRow() throws POIException {
		// do nothing
	}

	public void beforeRow() throws POIException {
		// do nothing
	}

}
