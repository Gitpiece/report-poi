package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IHeader;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.poi.operation.POIMarginOperation;
import com.cfcc.deptone.excel.poi.operation.POIOperation;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collection;
import java.util.Iterator;
import java.util.List;

/**
 * 构建表头。
 * @author WangHuanyu
 */
public class BuildHeader implements BuildStep {
	public String getName() {
		return "BuildHeader";
	}
	ISheet sheet;
	public void build(ISheet sheet) throws POIException {
		this.sheet = sheet;
		List<IHeader> allHeader = sheet.getAllHeader();
		for (Iterator<IHeader> iterator = allHeader.iterator(); iterator.hasNext();) {
			IHeader header = iterator.next();

			List<IPlaceHolder> holders = header.getAllPlaceHolder();

			Object value = header.getOriginalCellValue();
			for (IPlaceHolder hoder : holders) {
				value = POIExcelUtil.replaceMetadata(value.toString(), hoder.toPHString(), sheet.getMetadata().get(hoder.toArray()[1]));
			}

			//设置表头margin信息
			setMargin(header);

			sheet.setCellValue(header.getRow(), header.getColumn(), header, value);
			CellRangeAddress rangeAddress = header.getRangeAddress();
			if(rangeAddress != null){
				int offcolumn = rangeAddress.getLastColumn()-rangeAddress.getFirstColumn();
				rangeAddress.setFirstColumn(header.getColumn());
				rangeAddress.setLastColumn(header.getColumn() + offcolumn);
				header.setRangeAddress(rangeAddress);
//				sheet.setCellRangeNCellStyle(header);
				sheet.addMergedRegion(rangeAddress);
				sheet.setRegionCellStyle(header.getCellStyle(), rangeAddress);
			}
		}
	}

	private void setMargin(IHeader header) {
		int dataRightColumn = sheet.getDataRightColumn();
		int dataLeftColumn = sheet.getDataLeftColumn();
		Collection<POIOperation> poiOperations= header.getPoiOperation();
		if(poiOperations != null)
		for (POIOperation poiOperation:poiOperations){
			if(POIMarginOperation.OPERATION.equals(poiOperation.getOperation())){
				POIMarginOperation poiMarginOperation =(POIMarginOperation)poiOperation;
				if(ExcelConsts.MARGIN_CENTER.equals(poiMarginOperation.getMargintype())){
					CellRangeAddress cellRangeAddress = header.getRangeAddress();
					int column = dataLeftColumn + (dataRightColumn - dataLeftColumn)/2;
//					if(cellRangeAddress != null){
//						int rangecolumnlength = cellRangeAddress.getLastColumn()-cellRangeAddress.getFirstColumn();
//						//column -= rangecolumnlength/2;
//					}
					Integer number = poiMarginOperation.getNumber()==null?0:poiMarginOperation.getNumber();
					header.setCoordinate(header.getOriginalRow(),column+number);
				}else if(ExcelConsts.MARGIN_RIGHT.equals(poiMarginOperation.getMargintype())){
					int number = poiMarginOperation.getNumber();
					header.setCoordinate(header.getOriginalRow(),dataRightColumn-number);
				}else if(ExcelConsts.MARGIN_LEFT.equals(poiMarginOperation.getMargintype())){
					int number = poiMarginOperation.getNumber();
					header.setCoordinate(header.getOriginalRow(),dataLeftColumn+number);
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
