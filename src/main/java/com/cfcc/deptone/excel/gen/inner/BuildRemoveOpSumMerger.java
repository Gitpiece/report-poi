package com.cfcc.deptone.excel.gen.inner;

import org.apache.poi.ss.util.CellRangeAddress;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IOpSum;
import com.cfcc.deptone.excel.model.ISheet;

/**
 * 删除求合操作合并信息步骤
 * 
 * @author WangHuanyu
 */
public class BuildRemoveOpSumMerger implements BuildStep {
	public String getName() {
		return "BuildRemoveOpSumMerger";
	}

	public void build(ISheet sheet) throws POIException {

		for (IOpSum iOpSum : sheet.getAllOpSum()) {

			for (int i = sheet.getNumMergedRegions(); i > 0; i--) {
				CellRangeAddress rangeAddress = sheet.getMergedRegion(i - 1);

				if (iOpSum.getOriginalRow() >= rangeAddress.getFirstRow() && iOpSum.getOriginalRow() <= rangeAddress.getLastRow()
				    && iOpSum.getOriginalColumn() >= rangeAddress.getFirstColumn() && iOpSum.getOriginalColumn() <= rangeAddress.getLastColumn()) {

					iOpSum.setRangeAddress(rangeAddress);
					sheet.removeMergedRegion(i - 1);
				}
			}
		}
	}

	public void afterRow() throws POIException {
		//do nothing
	}

	public void beforeRow() throws POIException {
		//do nothing
	}

}
