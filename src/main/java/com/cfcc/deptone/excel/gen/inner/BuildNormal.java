package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.INormal;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.model.ISubTotal;
import com.cfcc.deptone.excel.poi.operation.POIMergeOperation;
import com.cfcc.deptone.excel.poi.operation.POIOperation;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Method;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 构建 sheet 按行写入数据步骤
 *
 * @author WangHuanyu
 */
public class BuildNormal implements BuildStep {

    private ISheet sheet;
    List<INormal> normals;
    Map<String, Method> methodcacheMap;
    /**
     * 小计集合
     */
    List<ISubTotal> subTotals ;
    /**
     * group normal
     */
    INormal groupNormal;
    /**
     * 数据对象
     */
    Object dataObj = null;

    /**
     * 数据写入的行数，从0开始。
     */
    int datawriterow = 0;

    public String getName() {
        return "BuildNormal";
    }

    public void build(ISheet sheet) throws POIException {
        this.sheet = sheet;
        List<?> dataList = sheet.getData();
        methodcacheMap = new HashMap<String, Method>(0);
        normals = sheet.getAllNormal();

        //按行写数据
        int datasize = dataList.size();
        //读取数据比实际数据多一条，为了增加beforeRow方法处理。
        for (int dataindex = 0; dataindex <= datasize; dataindex++) {

            dataObj = dataindex < datasize ? dataList.get(dataindex) : null;

            beforeRow();

            if (dataObj == null) continue;

            for (INormal normal : normals) {
                //
                Object value = getPropertyValue(methodcacheMap, datawriterow, dataObj, normal);
                normal.setCoordinate(normal.getOriginalRow() + datawriterow, normal.getOriginalColumn());
                normal.setCellStyle();
                sheet.setCellValue(normal.getOriginalRow() + datawriterow, normal.getOriginalColumn(), normal, value);
            }

            afterRow();
            datawriterow++;
        }

        //set offrow and offcolumn
        //因为循环时多读了一行，所以这里减2
        sheet.setDataOffRow(datawriterow - 1);
        sheet.setDataOffColumn(normals.size() - 1);

        doExtendOperation();

        //因为小计会占一行，所以如果报表既有小计，位移的行数减1，
        if(sheet.hasSubTotal()){
            sheet.setDataOffRow(datawriterow - 2);
        }

    }

    private void doExtendOperation() {
        // 在数据写完后进行列合并处理
        for (INormal iNormal : normals) {
            Collection<POIOperation> opCollection = iNormal.getPoiOperation();
            if (opCollection != null) {
                for (POIOperation poiOperation : opCollection) {
                    if (poiOperation.getOperation().endsWith(POIMergeOperation.OPERATION)) {
                        POIMergeOperation mo = (POIMergeOperation) poiOperation;
                        //列合并
                        if (ExcelConsts.VERTICAL.equals(mo.getMergeType())) {
                            doVMerge(iNormal);
                        }
                    }
                }
            }
        }
    }

    /**
     * 列合并
     *
     * @param iNormal
     */
    private void doVMerge(INormal iNormal) {
        /*
         * 首先保存第一行的值与行数，作为待比较数据。
		 * 循环：从第二行开始循环所有的数据行，循环时多读取一行，解决边界值为题。
		 * 当读取到数据与待比较数据一致时，继续读取下一行。
		 * 当读取到数据与待比较数据不一致，且相差行数等于1时，更新待比较数据与行数为当前数据与行数。
		 * 当读取到数据与待比较数据不一致，且相差行数大于1时，合并待比较数据行到当前数据前一行，更新待比较数据与行数为当前数据与行数。
		 */
        int originalrow = iNormal.getOriginalRow();
        int offrow = originalrow + sheet.getDataOffRow() + 2;
        //预读取一行数据，作为待比较数据
        String basevalue = sheet.getCellStringValue(originalrow, iNormal.getOriginalColumn());
        int baserow = originalrow;
        //循环
        for (int row = originalrow + 1; row < offrow; row++) {
            String comparevalue = sheet.getCellStringValue(row, iNormal.getOriginalColumn());
            //检查单元格内容是否匹配
            //内容部匹配并且行数间隔大于1，合并单元格
            if (!basevalue.equals(comparevalue)) {
                //更新比较字符串为当前
                if (row - baserow > 1) {
                    sheet.addMergedRegion(baserow, row - 1, iNormal.getColumn(), iNormal.getColumn());
                }
                basevalue = comparevalue;
                baserow = row;
            }
        }
    }

    /**
     * 获取属性值
     *
     * @param methodMap 方法缓存map
     * @param datarow   当前写入数据行数
     * @param dataObj   当前行数据
     * @param normal    普通报表接口
     * @return
     */
    private Object getPropertyValue(Map<String, Method> methodMap, int datarow, Object dataObj, INormal normal) {
        Object value;
        if (INormal.SN.equals(normal.getPropertyName())) {
            value = datarow + 1;
        } else {
            value = POIExcelUtil.getPropertyValue(dataObj, normal.getPropertyName(), methodMap);
        }
        return value;
    }

    public void afterRow() throws POIException {
        for (int i = 0; i < normals.size(); i++) {
            INormal normal = normals.get(i);
            //启用合并开关
            if (normal.isOpMerge()) {
                doHMerge(i, normal);
            }
        }
    }

    /*=========================================================
     *
     * 扩展选项
     *
     =========================================================*/
    private void doHMerge(int i, INormal normal) {
        String cvalue = normal.getCurrentCellStringValue();
        int num = normal.opMergeNumber();
        //检查单元格内容是否匹配，匹配的个数
        int j = 1;
        for (; j < num; j++) {
            INormal tmpNormal = normals.get(i + j);
            String tmpvalue = tmpNormal.getCurrentCellStringValue();
            if (!cvalue.equals(tmpvalue)) {
                break;
            }
        }
        if (j == num) {
            //如果定义的合并单元格个数和实际匹配的一致，合并单元格。
            INormal merge2Normal = normals.get(i + num - 1);
            //注释 CellRangeAddress range = new CellRangeAddress(normal.getRow(),merge2Normal.getRow() , normal.getColumn(), merge2Normal.getColumn());
            sheet.addMergedRegion(normal.getRow(), merge2Normal.getRow(), normal.getColumn(), merge2Normal.getColumn());
        }
    }


    public void beforeRow() throws POIException {
        writesubtotal();

    }

    /**
     * 写入数据行数
     */
    int subtotalstartrow = 0;

    /**
     * 写入小计
     */
    private void writesubtotal() {
        //初始化小计信息
        initsubtotal();

        //如果有小计，增加小计计算
        if (sheet.hasSubTotal()) {
            //记录写入数据行信息

            //判断是否分组结束，如果分组结束，计算小计
            if(datawriterow == 0) return;
            Object precurrentvalue = sheet.getCellStringValue(this.groupNormal.getOriginalRow()+datawriterow - 1,this.groupNormal.getOriginalColumn());//getPropertyValue(methodcacheMap, datawriterow - 1, dataObj, this.groupNormal);

            //如果当前数据对象为Null，说明已经写入到最后一行
            // 或者当前值不等于上一行的值
            if (dataObj == null || !precurrentvalue.equals(getPropertyValue(methodcacheMap, datawriterow, dataObj, this.groupNormal)))
            {
                //插入小计
//                System.out.println(precurrentvalue);
//                System.out.println(this.datawriterow);
                //先按照normal写数据，直到分组字段开始写小计信息。
                for (INormal normal : this.normals) {
                    Object value="";
                    if (normal.hasGroup()) {
                        value = precurrentvalue;
                    }else{
                        value = getPropertyValue(methodcacheMap, datawriterow, dataObj, normal);
                    }
                    //
                    normal.setCoordinate(normal.getOriginalRow() + datawriterow, normal.getOriginalColumn());
                    normal.setCellStyle();
                    sheet.setCellValue(normal.getOriginalRow() + datawriterow, normal.getOriginalColumn(), normal, value);
                    if (normal.hasGroup()) {
                        break;
                    }
                }
                //小计
                for (ISubTotal iSubTotal : this.subTotals) {
                    iSubTotal.setCoordinate(this.groupNormal.getOriginalRow()+datawriterow,iSubTotal.getOriginalColumn());
                    iSubTotal.setCellStyle();
                    if(ExcelConsts.STATIC.equals(iSubTotal.getType())){
                        sheet.setCellValue(iSubTotal.getRow(), iSubTotal.getColumn(), iSubTotal, iSubTotal.getValue());
                    }else  if(ExcelConsts.CALC.equals(iSubTotal.getType())){
                        Cell cell = sheet.getCell(iSubTotal.getRow(), iSubTotal.getColumn());
                        cell.setCellFormula(POIExcelUtil.formatAsSumString(groupNormal.getOriginalRow()+subtotalstartrow, iSubTotal.getColumn(), groupNormal.getOriginalRow()+datawriterow-1, iSubTotal.getColumn()));
                    }
                }

                //因为增加了小计，数据行+1、更新subtotalstartrow
                datawriterow++;
                subtotalstartrow = datawriterow;
            }
        }
    }


    /**
     * 初始化小计信息
     */
    private void initsubtotal() {
        if (sheet.hasSubTotal() && groupNormal == null) {
            //记录分组单元格信息
            for (INormal iNormal : this.normals) {
                if (iNormal.hasGroup()) {
                    groupNormal = iNormal;
                }
            }
            subTotals = sheet.getAllSubtotals();

            //整理小计集合。正常情况，normal会比subtotal至少长1。
            //把subtotal整理成与normal长度一致。
//            int normalindex=0,subtotalindex=0;
//            for (INormal normal:this.normals){
//                //如果
//                if (INormal.SN.equals(normal.getPropertyName())) {
//
//                }
//
//                normalindex++;
//            }
        }
    }

}
