package com.cfcc.deptone.excel.poi.operation;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.util.StringManager;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import java.util.ArrayList;
import java.util.Collection;

/**
 * poi operation解析类。
 *
 * <p>operation是用来控制报表展示的一些扩展属性，以批注的方式添加到单元格上。</p>
 */
public class POIOperationParser {
	private static final Log LOGGER = LogFactory.getLog(POIOperationParser.class);

	StringManager sm = StringManager.getManager(Constants.Package);
	/**
	 * 解析 operation
	 * @param operationString
	 * @return
	 */
	public Collection<POIOperation> parseCacheAnnotations(String operationString) throws POIException {
		
		Collection<POIOperation> ops = null;
		String[] operationArray = operationString.split("\n");
		for (String opString : operationArray) {
			String[] operationStringArray = opString.split(" ");
			if(POIOrderOperation.OPERATION.equals(operationStringArray[0])){
				ops = lazyInit(ops);
				ops.add(parseOrderOperation(operationStringArray));
			}else if(SpreadOperation.OPERATION.equals(operationStringArray[0])){
				ops = lazyInit(ops);
				ops.add(parseSpreadOperation(operationStringArray));
			}else if(POIMergeOperation.OPERATION.equals(operationStringArray[0])){
				ops = lazyInit(ops);
				ops.add(parseMergeOperation(operationStringArray));
			}else if(POIGroupOperation.OPERATION.equals(operationStringArray[0])){
				ops = lazyInit(ops);
				ops.add(parseGroupOperation(operationStringArray));
			}else if(POIMarginOperation.OPERATION.equals(operationStringArray[0])){
				ops = lazyInit(ops);
				ops.add(parseMarginOperation(operationStringArray));
			}else{
				LOGGER.warn("unknow poi operation :" + opString);
				throw new POIException(sm.getString(sm.getString("poi.operation.unknow"),opString));
			}
		}
		return ops;
	}

	private POIOperation parseMarginOperation(String[] operationStringArray) {
		POIMarginOperation poiMarginOperation = new POIMarginOperation();
		poiMarginOperation.setOperation(operationStringArray[0]);
		poiMarginOperation.setMargintype(operationStringArray[1]);
		if(operationStringArray.length >2){
			poiMarginOperation.setNumber(Integer.parseInt(operationStringArray[2]));
		}
		return poiMarginOperation;
	}


	private POIOperation parseGroupOperation(String[] operationStringArray) {
		POIGroupOperation groupOperation = new POIGroupOperation();
		groupOperation.setOperation(operationStringArray[0]);
		return groupOperation;
	}

	private POIOperation parseSpreadOperation(String[] opStrArr) {
		SpreadOperation operation= new SpreadOperation();
		operation.setOperation(opStrArr[0]);
		operation.setSpreadType(opStrArr.length>1?opStrArr[1]:SpreadOperation.SPREAD_TYPE_ALL);
		return operation;
	}

	private POIOperation parseOrderOperation(String[] operationStringArray) {
		POIOrderOperation operation= new POIOrderOperation();
		operation.setOperation(operationStringArray[0]);
		operation.setOrderType(operationStringArray[1]);
		return operation;
	}
	
	private POIOperation parseMergeOperation(String[] operationStringArray) {
		POIMergeOperation operation= new POIMergeOperation();
		operation.setOperation(operationStringArray[0]);
		operation.setMergeType(operationStringArray[1]);
		return operation;
	}
	
	private <T extends POIOperation> Collection<POIOperation> lazyInit(Collection<POIOperation> ops) {
		return ops != null ? ops : new ArrayList<POIOperation>(1);
	}
}
