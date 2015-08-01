package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IHeader;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.poi.operation.POIMarginOperation;
import com.cfcc.deptone.excel.poi.operation.POIOperation;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collection;
import java.util.Iterator;
import java.util.List;

/**
 * 构建表头。
 *
 * @author WangHuanyu
 */
public class BuildHeader implements BuildStep {
    private Log logger = LogFactory.getLog(BuildHeader.class);

    public String getName() {
        return "BuildHeader";
    }

    ISheet sheet;

    public void build(ISheet sheet) {
        if (logger.isInfoEnabled()) {
            logger.info(this.getName());
        }
        this.sheet = sheet;
        List<IHeader> allHeader = sheet.getAllHeader();
        for (Iterator<IHeader> iterator = allHeader.iterator(); iterator.hasNext(); ) {
            IHeader header = iterator.next();

            List<IPlaceHolder> holders = header.getAllPlaceHolder();

            Object value = header.getOriginalCellValue();
            for (IPlaceHolder hoder : holders) {
                value = POIExcelUtil.replaceMetadata(value.toString(), hoder.toPHString(), sheet.getMetadata().get(hoder.toArray()[1]));
            }

            //设置表头margin信息
            setMargin(header);

            //try {
                sheet.setCellValue(header.getRow(), header.getColumn(), header, value);
            //} catch (IllegalArgumentException ie) {}
            CellRangeAddress rangeAddress = header.getRangeAddress();
            if (rangeAddress != null) {
                int offcolumn = rangeAddress.getLastColumn() - rangeAddress.getFirstColumn();
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
        Collection<POIOperation> poiOperations = header.getMarginPoiOperation();
        if (poiOperations != null && poiOperations.size() > 0)
            for (POIOperation poiOperation : poiOperations) {
                POIMarginOperation poiMarginOperation = (POIMarginOperation) poiOperation;
                if (ExcelConsts.MARGIN_CENTER.equals(poiMarginOperation.getMargintype())) {
                    int column = dataLeftColumn + (dataRightColumn - dataLeftColumn) / 2;
                    Integer number = poiMarginOperation.getNumber() == null ? 0 : poiMarginOperation.getNumber();
                    header.setCoordinate(header.getOriginalRow(), column + number);
                } else if (ExcelConsts.MARGIN_RIGHT.equals(poiMarginOperation.getMargintype())) {
                    int number = poiMarginOperation.getNumber();
                    header.setCoordinate(header.getOriginalRow(), dataRightColumn - number);
                } else if (ExcelConsts.MARGIN_LEFT.equals(poiMarginOperation.getMargintype())) {
                    int number = poiMarginOperation.getNumber();
                    header.setCoordinate(header.getOriginalRow(), dataLeftColumn + number);
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
