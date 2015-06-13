package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.*;
import com.cfcc.deptone.excel.poi.POIFactory;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.poi.ss.usermodel.Cell;

import java.util.List;

/**
 * 构建 sheet 交叉求和步骤,因该在写入数据之后执行本步骤.
 * <br />
 * 通过数据已经写入的情况，构建sum占位符，交给{@link BuildOpSum}完成求和单元格的构建。
 * 
 * @author wanghuanyu
 */
public class BuildCrossTabSum implements BuildStep {

	public String getName() {
		return "BuildCrossTabSum";
	}

	/**
	 * 把交叉报表的求和计算转化为普通报表的求和。
	 */
	public void build(ISheet sheet) throws POIException {
		List<ICrossSum> allCrossSum = sheet.getAllCrossSum();
		List<IOpSum> allOpsum = sheet.getAllOpSum();
		boolean vertical=false,horizontal = false;
		for (int j = 0; j < allCrossSum.size(); j++) {
			ICrossSum crossTab = allCrossSum.get(j);
			
			//求和是横向还是纵向，现在只支持纵向
			if(crossTab.isExpand() || crossTab.isSpread()){
				if(ExcelConsts.VERTICAL.equals(crossTab.getCoordinate())){
					vertical=true;
					for(int i = 0; i <= sheet.getDataOffColumn(); i++){
						String[] arr = new String[3];
						arr[0] = "sum";
						arr[1] = crossTab.getCoordinate();
						arr[2] = "crossTabSum_"+crossTab.getCoordinate()+"_"+i;
						
						IOpSum opSum = POIFactory.buildIOpSum(sheet, crossTab.getCell(), arr);
						//因为交叉报表时动态的，所以原始列需要做相应的移动
						opSum.setOriginalColumn(opSum.getOriginalColumn()+i);
						allOpsum.add(opSum);
					}
				} else if(ExcelConsts.HORIZONTAL.equals(crossTab.getCoordinate())){
					horizontal = true;
					for(int i = 0; i <= sheet.getDataOffRow(); i++){
						String[] arr = new String[3];
						arr[0] = "sum";
						arr[1] = crossTab.getCoordinate();
						arr[2] = "crossTabSum_"+crossTab.getCoordinate()+"_"+i;
						
						IOpSum opSum = POIFactory.buildIOpSum(sheet, crossTab.getCell(), arr);
						//因为交叉报表时动态的，所以原始列需要做相应的移动
						opSum.getOriginalRow(opSum.getOriginalRow()+i);
						allOpsum.add(opSum);
					}
				}
			} else {
				String[] arr = new String[3];
				arr[0] = "sum";
				arr[1] = crossTab.getCoordinate();
				arr[2] = "crossTabSum_"+ ExcelConsts.VERTICAL;
				
				IOpSum opSum = POIFactory.buildIOpSum(sheet, crossTab.getCell(), arr);
				allOpsum.add(opSum);
			}
		}

		//如果行列都有求和，增加右下角的空白单元格，定义常量
		if(allCrossSum.size() >0 && vertical && horizontal){
			List<ICrossTab> iCrossTabs = POIExcelUtil.getAllDataCross(sheet.getAllCrossTab());
			ICrossTab iCrossTab = iCrossTabs.get(0);
			//因为增加了合集列，所以在位移的基础上加一
			Cell cell = sheet.getCell(iCrossTab.getOriginalRow()+sheet.getDataOffRow()+1,iCrossTab.getOriginalColumn()+sheet.getDataOffColumn()+1);
			cell.setCellStyle(iCrossTab.getCellStyle());
			String[] arr = new String[4];
			arr[0] = CellType.CONST.getName();
			arr[1] = ExcelConsts.FIX;
			arr[2] = ExcelConsts.STATIC;
			arr[3] = "";
			sheet.getAllConst().add(POIFactory.buildIConst(sheet, cell, arr));
		}
	}

	public void afterRow() throws POIException {
		//do nothing
	}

	public void beforeRow() throws POIException {
		//do nothing
	}
}
