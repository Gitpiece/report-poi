package com.cfcc.deptone.excel.gen.inner;

import org.apache.poi.ss.util.CellRangeAddress;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IFooter;
import com.cfcc.deptone.excel.model.ISheet;

/**
 * 删除表底合并信息步骤
 * 
 * @author WangHuanyu
 */
public class BuildRemoveFooterMerger implements BuildStep {
	public String getName() {
		return "BuildRemoveFooterMerger";
	}

	public void build(ISheet sheet) throws POIException {

		for (IFooter footer : sheet.getAllFooter()) {

			for (int i = sheet.getNumMergedRegions(); i > 0; i--) {
				CellRangeAddress rangeAddress = sheet.getMergedRegion(i - 1);

				if (footer.getOriginalRow() >= rangeAddress.getFirstRow() && footer.getOriginalRow() <= rangeAddress.getLastRow()
				    && footer.getOriginalColumn() >= rangeAddress.getFirstColumn() && footer.getOriginalColumn() <= rangeAddress.getLastColumn()) {

					footer.setRangeAddress(rangeAddress);
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
