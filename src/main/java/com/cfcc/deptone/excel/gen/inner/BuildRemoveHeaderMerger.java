package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IHeader;
import com.cfcc.deptone.excel.model.ISheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 删除表底合并信息步骤
 * 
 * @author WangHuanyu
 */
public class BuildRemoveHeaderMerger implements BuildStep {
	public String getName() {
		return "BuildRemoveFooterMerger";
	}

	public void build(ISheet sheet) throws POIException {

		for (IHeader iHeader : sheet.getAllHeader()) {

			for (int i = sheet.getNumMergedRegions(); i > 0; i--) {
				CellRangeAddress rangeAddress = sheet.getMergedRegion(i - 1);

				if (iHeader.getOriginalRow() >= rangeAddress.getFirstRow() && iHeader.getOriginalRow() <= rangeAddress.getLastRow()
				    && iHeader.getOriginalColumn() >= rangeAddress.getFirstColumn() && iHeader.getOriginalColumn() <= rangeAddress.getLastColumn()) {

					iHeader.setRangeAddress(rangeAddress);
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
