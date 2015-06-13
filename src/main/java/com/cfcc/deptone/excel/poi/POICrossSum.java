package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ICrossSum;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.poi.operation.POIOperation;
import com.cfcc.deptone.excel.poi.operation.SpreadOperation;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Collection;
import java.util.StringTokenizer;

/**
 * 交叉报表求和。
 *
 * @author wanghuanyu
 */
public class POICrossSum extends POICellObject implements ICrossSum {

    private String coordinate = null;
    private boolean isExpand = false;
    //
    private boolean spread = false;//传播性
    private String spreadType = null;

    /**
     * @param poiSheet sheet object
     * @param cell     单元格对象
     * @param arr      eg. #{crosstab.sum.v.[expand]}
     */
    protected POICrossSum(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
        super(poiSheet, cell);

        coordinate = arr[2];
        //FIXME WHY 兼容1.2.1之前版本
        if (arr.length > 3) {
            String tmp = arr[3].replace("[", "");
            String option = tmp.replace("]", "");
            StringTokenizer st = new StringTokenizer(option);
            while (st.hasMoreTokens()) {
                if (st.nextToken().equalsIgnoreCase(ICrossSum.EXPAND)) {
                    isExpand = true;
                }
            }
        }
    }

    @Override
    public void afterInit() {
        super.afterInit();
        //处理扩展参数
        Collection<POIOperation> poiOperationCollection = this.getPoiOperation();
        if (poiOperationCollection != null) {
            for (POIOperation poiOperation : poiOperationCollection) {
                if (SpreadOperation.OPERATION.equals(poiOperation.getOperation())) {
                    this.spread = true;
                    this.spreadType = ((SpreadOperation) poiOperation).getSpreadType();
                }
            }
        }
    }

    public String getCoordinate() {
        return coordinate;
    }

    /**
     * @deprecated replaced by {@link #isSpread()}
     */
    @Deprecated
    public boolean isExpand() {
        return isExpand;
    }

    public boolean isSpread() {
        return spread;
    }

    public String getSpreadType() {
        return spreadType;
    }
}
