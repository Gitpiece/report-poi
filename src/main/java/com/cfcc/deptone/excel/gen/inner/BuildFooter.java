package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IFooter;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.poi.operation.POIMarginOperation;
import com.cfcc.deptone.excel.poi.operation.POIOperation;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collection;
import java.util.List;

/**
 * 构建 sheet 表底步骤
 * 
 * @author WangHuanyu
 */
public class BuildFooter implements BuildStep {
	public String getName() {
		return "BuildFooter";
	}
	ISheet sheet;
	public void build(ISheet sheet) throws POIException {
		this.sheet = sheet;
		List<IFooter> allFooter = sheet.getAllFooter();
		int offRow = 0;
		for (IFooter footer : allFooter) {
			if (offRow == 0) {
				offRow = footer.getOriginalRow();
			}
			
			Object value = footer.getOriginalCellValue();
			for (IPlaceHolder hoder : footer.getAllPlaceHolder()) {
				value = POIExcelUtil.replaceMetadata(value.toString(), hoder.toPHString(), sheet.getMetadata().get(hoder.toArray()[1]));
			}


			setMargin(footer);

			sheet.setCellValue(footer.getRow() + sheet.getDataOffRow(), footer.getColumn(),footer, value);

			sheet.setCellRangeNCellStyle(footer);

			//
			offRow = (footer.getOriginalRow() - offRow) > 0 ? footer.getOriginalRow() - offRow : offRow;
		}

		//
		sheet.setDataOffRow(sheet.getDataOffRow() + offRow);
	}

	private void setMargin(IFooter iFooter) {
		int dataRightColumn = sheet.getDataRightColumn();
		int dataLeftColumn = sheet.getDataLeftColumn();
		Collection<POIOperation> poiOperations= iFooter.getPoiOperation();
		if(poiOperations != null)
			for (POIOperation poiOperation:poiOperations){
				if(POIMarginOperation.OPERATION.equals(poiOperation.getOperation())){
					POIMarginOperation poiMarginOperation =(POIMarginOperation)poiOperation;
					if(ExcelConsts.MARGIN_CENTER.equals(poiMarginOperation.getMargintype())){
						CellRangeAddress cellRangeAddress = iFooter.getRangeAddress();
						int column = dataLeftColumn + (dataRightColumn - dataLeftColumn)/2;
//					if(cellRangeAddress != null){
//						int rangecolumnlength = cellRangeAddress.getLastColumn()-cellRangeAddress.getFirstColumn();
//						//column -= rangecolumnlength/2;
//					}
						Integer number = poiMarginOperation.getNumber()==null?0:poiMarginOperation.getNumber();
						iFooter.setCoordinate(iFooter.getOriginalRow(),column+number);
					}else if(ExcelConsts.MARGIN_RIGHT.equals(poiMarginOperation.getMargintype())){
						int number = poiMarginOperation.getNumber();
						iFooter.setCoordinate(iFooter.getOriginalRow(),dataRightColumn-number);
					}else if(ExcelConsts.MARGIN_LEFT.equals(poiMarginOperation.getMargintype())){
						int number = poiMarginOperation.getNumber();
						iFooter.setCoordinate(iFooter.getOriginalRow(),dataLeftColumn+number);
					}
				}
			}
	}


	public void afterRow() throws POIException {
		// do nothking
	}

	public void beforeRow() throws POIException {
		// do nothking
	}

}
