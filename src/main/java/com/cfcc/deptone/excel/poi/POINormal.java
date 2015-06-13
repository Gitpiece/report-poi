package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.INormal;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.poi.operation.POIGroupOperation;
import com.cfcc.deptone.excel.poi.operation.POIOperation;
import org.apache.poi.ss.usermodel.Cell;

import java.util.List;

public class POINormal extends POICellObject implements INormal {

	boolean merge = false;
	int mergeNumber = 0;
	boolean group = false;
	protected POINormal(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
		super(poiSheet, cell);
		this.propertyName = arr[1];
		
		if(arr.length>=3){
			String optional = arr[2];
			optional = optional.replace("[", "").replace("]", "");
			String[] optionalArr = optional.split(" ");
			for(int i=0;i<optionalArr.length;i++){
				if(OPTIONAL_MERGE.equals(optionalArr[i])){
					merge = true;
					mergeNumber = Integer.parseInt(optionalArr[++i]);
				}
			}
		}

		completion();
	}

	protected void completion() {
		if(this.getPoiOperation() != null)
		for (POIOperation poiOperation :this.getPoiOperation()){
			if(POIGroupOperation.OPERATION.equals(poiOperation.getOperation())){
				this.group = true;
			}
		}
	}

	public String getPropertyName() {
		List<IPlaceHolder> list = this.getAllPlaceHolder();
		return list.get(0).toArray()[1];
	}
	/*=====================================================
	 * 
	 * 常规单元格借口参数
	 * 
	 =====================================================*/
	public boolean isOpMerge() {
		return merge;
	}

	public int opMergeNumber() {
		return mergeNumber;
	}


	public boolean hasGroup(){
		return this.group;
	}

}
