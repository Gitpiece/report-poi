package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IConst;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import com.cfcc.deptone.excel.util.StringManager;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * 构建 sheet 常量步骤
 * 
 * @author WangHuanyu
 */
public class BuildConst implements BuildStep {

    StringManager sm = StringManager.getManager("com.cfcc.deptone.excel.gen.inner");
	public String getName() {
		return "BuildConst";
	}

	public void build(ISheet sheet) throws POIException {
		List<IConst> allConst = sheet.getAllConst();
		for (IConst iConst : allConst) {

			//得到常量值
			Object value = "";
			if (ExcelConsts.STATIC.equals(iConst.getType())) {
				value = iConst.getPropertyName();
			} else if (ExcelConsts.PARAM.equals(iConst.getType())) {
				List<IPlaceHolder> holders = iConst.getAllPlaceHolder();
				
				value = iConst.getOriginalCellValue();
				for (IPlaceHolder hoder : holders) {
					value = POIExcelUtil.replaceMetadata(value.toString(), hoder.toPHString(), sheet.getMetadata().get(hoder.toArray()[3]));
				}
			}

			int row = iConst.getOriginalRow(), column = iConst.getOriginalColumn();
			if (ExcelConsts.VERTICAL.equals(iConst.getCoordinate())) {
				row += sheet.getDataOffRow();
			} else if (ExcelConsts.HORIZONTAL.equals(iConst.getCoordinate())) {
				column += sheet.getDataOffColumn();
			} else if(!ExcelConsts.FIX.equals(iConst.getCoordinate())){
				// 如果不是FIX，说明是不支持的占位符
                throw new POIException(sm.getString("poi.placeholder.notsupport",iConst.getOriginalCellValue()));
			}
			
			iConst.setCoordinate(row, column);
			iConst.setValue(value);
			
			setCellRange(sheet, iConst);
			
		}
	}

	private void setCellRange(ISheet sheet, IConst iConst) {
	    //设置合并单元格
	    CellRangeAddress cellRange = iConst.getRangeAddress();
	    if (cellRange != null) {
	    	if (ExcelConsts.VERTICAL.equals(iConst.getCoordinate())) {
	    		cellRange.setFirstRow(cellRange.getFirstRow() + sheet.getDataOffRow());
	    		cellRange.setLastRow(cellRange.getLastRow() + sheet.getDataOffRow());
	    	}else if (ExcelConsts.HORIZONTAL.equals(iConst.getCoordinate())) {
	    		cellRange.setFirstColumn(cellRange.getFirstColumn() + sheet.getDataOffColumn());
	    		cellRange.setLastColumn(cellRange.getLastColumn() + sheet.getDataOffColumn());
	    	}
	    	sheet.addMergedRegion(cellRange);
	    	//设置单元格格式
	    	sheet.setRegionCellStyle(iConst.getCellStyle(), cellRange);
	    }else{
	    	//设置单元格格式
	    	iConst.setCellStyle();
	    }
    }

	public void afterRow() throws POIException {
		//do nothking
	}

	public void beforeRow() throws POIException {
		//do nothking
	}

}
