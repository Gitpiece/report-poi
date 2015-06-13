package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.*;
import com.cfcc.deptone.excel.model.support.InitBean;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public class POIFactory {

    private ISheet poiSheet;

    public POIFactory(POISheet poiSheet) {
        this.poiSheet = poiSheet;
    }

    /**
     * 把poi cell包装成自定义单元格类型
     *
     * @param cell
     * @return
     */
    public POICellObject wrap(Cell cell) throws POIException {

        POICellObject result = null;
        if (Cell.CELL_TYPE_STRING == cell.getCellType() && !POIExcelUtil.isEmpty(cell.getStringCellValue())) {

            List<IPlaceHolder> list = POIExcelUtil.toPlaceHolders(cell.getStringCellValue());

            if (!list.isEmpty()) {

                String[] arr = list.get(0).toArray();

                if (CellType.HEADER.getName().equals(arr[0])) {
                    result = new POIHeader(poiSheet, cell, arr);
                } else if (CellType.NORMAL.getName().equals(arr[0])) {
                    result = new POINormal(poiSheet, cell, arr);
                } else if (CellType.FOOTER.getName().equals(arr[0])) {
                    result = new POIFooter(poiSheet, cell, arr);
                } else if (CellType.CROSSTAB.getName().equals(arr[0])) {
                    if (CellType.CROSSSUM.getName().equals(arr[1])) {
                        result = new POICrossSum(poiSheet, cell, arr);
                    } else {
                        result = new POICrossTab(poiSheet, cell, arr);
                    }
                } else if (CellType.SUM.getName().equals(arr[0])) {
                    result = new POIOpSum(poiSheet, cell, arr);
                } else if (CellType.CONST.getName().equals(arr[0])) {
                    result = new POIConst(poiSheet, cell, arr);
                } else if (CellType.PICTURE.getName().equals(arr[0])) {
                    result = new POIPicture(poiSheet, cell, arr);
                } else if (CellType.SUBTOTAL.getName().equals(arr[0])) {
                    result = new POISubTotal(poiSheet, cell, arr);
                } else { // undefined cell type
                    result = new POICellObject(poiSheet, cell);
                }
            }
        } else if (Cell.CELL_TYPE_FORMULA == cell.getCellType()) {
            result = new POIFormula(poiSheet, cell);
        }

        //初始化后操作
        if (result instanceof InitBean) {
            ((InitBean) result).afterInit();
        }
        return result;
    }

    public static IOpSum buildIOpSum(ISheet poiSheet, Cell cell, String[] arr) throws POIException {
        return new POIOpSum(poiSheet, cell, arr, false);
    }

    public static ISheet createSheet(Workbook workbook, Sheet sheet) {
        return new POISheet(workbook, sheet);
    }

    public static IPlaceHolder createPlaceHolder(String placeHolderStr) {
        return new POIPlaceHolder(placeHolderStr);
    }

    public static IConst buildIConst(ISheet sheet, Cell cell, String[] arr) throws POIException {
        return new POIConst(sheet, cell, arr, false);
    }
}
