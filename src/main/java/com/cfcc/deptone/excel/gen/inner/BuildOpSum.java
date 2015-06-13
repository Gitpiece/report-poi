package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IOpSum;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 构建 sheet 求合操作步骤
 * 
 * @author WangHuanyu
 */
public class BuildOpSum implements BuildStep {
	public String getName() {
		return "BuildOpSum";
	}

	public void build(ISheet sheet) throws POIException {
		List<IOpSum> allOpSum = sheet.getAllOpSum();
		Map<String, IOpSum> opSumCache = new HashMap<String, IOpSum>();

		for (IOpSum iOpSum : allOpSum) {
			CellRangeAddress cellRange = iOpSum.getRangeAddress();
			int row = iOpSum.getOriginalRow(), column = iOpSum.getOriginalColumn();
			int rangeCol = 0;
			Cell cell ;
			if (ExcelConsts.VERTICAL.equals(iOpSum.getCoordinate())) {
				row += sheet.getDataOffRow();
			} else if (ExcelConsts.HORIZONTAL.equals(iOpSum.getCoordinate())) {
				column += sheet.getDataOffColumn();
			}
			//获取单元格
			cell = sheet.getCell(row, column);
			//增加合并单元格坐标偏移
			if (cellRange != null) {
				rangeCol = cellRange.getLastColumn()-cellRange.getFirstColumn();
			}

			iOpSum.setCoordinate(row, column);
			opSumCache.put(iOpSum.getName(), iOpSum);

			//设置单元格类型
			if (!ExcelConsts.CELL_FORMATE_TEXT.equals(iOpSum.getDataFormat())) {
				cell.setCellType(Cell.CELL_TYPE_BLANK);
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			}
			
			if (iOpSum.hasMember()) {
				cell.setCellFormula(POIExcelUtil.format4Member(opSumCache, iOpSum.getMembers()));
			} else {
				int firstrow=0,firstcolumn=0,lastrow=0,lastcolumn=0;
				if(ExcelConsts.VERTICAL.equals(iOpSum.getCoordinate())){
					firstrow=iOpSum.getOriginalRow() - 1;
					firstcolumn= iOpSum.getOriginalColumn();
					lastrow=row - 1;
					lastcolumn=column + rangeCol;
					if(sheet.hasSubTotal()){
						firstrow--;
					}
				} else if (ExcelConsts.HORIZONTAL.equals(iOpSum.getCoordinate())) {
					firstrow=iOpSum.getOriginalRow();
					firstcolumn= iOpSum.getOriginalColumn() - 1;
					lastrow=row - row;
					lastcolumn=column + rangeCol - 1;
				}

				cell.setCellFormula(POIExcelUtil.formatAsSumStringHalf(firstrow, firstcolumn, lastrow, lastcolumn,sheet.hasSubTotal()));
			}
			
			sheet.setCellRangeNCellStyle(iOpSum);
		}

	}

	public void afterRow() throws POIException {
		//do nothking
	}

	public void beforeRow() throws POIException {
		//do nothking
	}

}
