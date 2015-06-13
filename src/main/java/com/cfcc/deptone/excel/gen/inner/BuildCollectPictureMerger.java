package com.cfcc.deptone.excel.gen.inner;

import org.apache.poi.ss.util.CellRangeAddress;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IPicture;
import com.cfcc.deptone.excel.model.ISheet;

/**
 * 删除表底合并信息步骤
 * 
 * @author WangHuanyu
 */
public class BuildCollectPictureMerger implements BuildStep {
	public String getName() {
		return "BuildCollectPictureMerger";
	}

	public void build(ISheet sheet) throws POIException {

		for (IPicture pic : sheet.getAllPicture()) {

			for (int i = sheet.getNumMergedRegions(); i > 0; i--) {
				CellRangeAddress rangeAddress = sheet.getMergedRegion(i - 1);

				if (pic.getOriginalRow() >= rangeAddress.getFirstRow() && pic.getOriginalRow() <= rangeAddress.getLastRow()
				    && pic.getOriginalColumn() >= rangeAddress.getFirstColumn() && pic.getOriginalColumn() <= rangeAddress.getLastColumn()) {
					
					pic.setRangeAddress(rangeAddress);
				}
			}
		}
	}

	public void afterRow() throws POIException {
		//do nothking
	}

	public void beforeRow() throws POIException {
		//do nothking
	}

}
