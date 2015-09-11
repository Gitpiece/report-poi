package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.gen.inner.BuildStep;
import com.cfcc.deptone.excel.model.*;
import com.cfcc.deptone.excel.util.Assert;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import com.cfcc.deptone.excel.util.StringManager;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.math.BigDecimal;
import java.util.*;

/**
 * excel sheet实现类。 <br />
 * 本类的主流程如下：
 * <p/>
 * <pre>
 * 	1，load方法加载所有模板定义信息，分类保存。如：header-&gt;headerCells,footer-&gt;footerCells。
 * 	2，scop方法根据加载的模板信息判断，确定sheet的类型。
 * 	3，调用build，按定义好的步骤写入数据。
 * </pre>
 *
 * @author WangHuanyu
 */
public class POISheet implements ISheet {
    private Log logger = LogFactory.getLog(POISheet.class);
    protected StringManager sm = StringManager.getManager("com.cfcc.deptone.excel.gen.inner");
    private boolean loaded = false;
    private Workbook workbook = null;
    protected Sheet sheet = null;
    protected CreationHelper creationHelper;

    private SheetType sheetType = null;
    protected POIFactory factory;

    private List<ICellObject> allCells = new ArrayList<ICellObject>(0);// 所有已定义cell，除other外。
    private List<IHeader> headerCells = new ArrayList<IHeader>(0);
    private List<IFooter> footerCells = new ArrayList<IFooter>(0);
    private List<IConst> constCells = new ArrayList<IConst>(0);
    private List<INormal> normalCells = new ArrayList<INormal>(0);
    private List<ISubTotal> subtotalCells = new ArrayList<ISubTotal>(0);
    private List<ICrossTab> crossTabCells = new ArrayList<ICrossTab>(0);
    private List<ICrossSum> crossSumCells = new ArrayList<ICrossSum>(0);
    private List<IOpSum> opSumCells = new ArrayList<IOpSum>(0);
    private List<IPicture> picCells = new ArrayList<IPicture>(0);
    private List<IFormula> formulaCells = new ArrayList<IFormula>(0);
    private List<ICellObject> others = new ArrayList<ICellObject>(0);

    private Map<String, Object> metaData = null;
    /**
     * bigDecimal 转化为string的样式缓存，key为定义的占位符。
     */
    private Map<String,CellStyle> bigdecimald2string = new HashMap<String, CellStyle>(10);
    private List<?> data = null;

    //	int startRow = 0;
//	int endRow = 0;
    /**
     * 数据写入行数，包括小计。
     */
    int dataOffRow = 0;
    /**
     * 数据写入列数。
     */
    private int dataOffColumn;
    /**
     * 是否包含小计标记，目前只支持normal报表。
     */
    private boolean subtotal;

    protected POISheet(Workbook workbook, Sheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.factory = new POIFactory(this);

        creationHelper = workbook.getCreationHelper();
    }

    public void load(Map<String, Object> metaData, List<?> data)
            throws POIException {
        this.metaData = metaData;
        this.data = data;

        int firstRowNum = this.sheet.getFirstRowNum();
        int lashRowNum = this.sheet.getLastRowNum();
        for (int row = firstRowNum; row <= lashRowNum; row++) {

            if (sheet.getRow(row) == null) {
                continue;
            }

            for (int j = sheet.getRow(row).getFirstCellNum(); j < sheet.getRow(
                    row).getLastCellNum(); j++) {

                Cell cell = sheet.getRow(row).getCell(j);
                if (cell == null) {
                    continue;
                }

                POICellObject cellObj = factory.wrap(cell);

                if (cellObj instanceof IHeader) {
                    headerCells.add((IHeader) cellObj);
                } else if (cellObj instanceof IFooter) {
                    footerCells.add((IFooter) cellObj);
                } else if (cellObj instanceof INormal) {
                    normalCells.add((INormal) cellObj);
                } else if (cellObj instanceof ICrossTab) {
                    crossTabCells.add((ICrossTab) cellObj);
                } else if (cellObj instanceof IOpSum) {
                    opSumCells.add((IOpSum) cellObj);
                } else if (cellObj instanceof IConst) {
                    constCells.add((IConst) cellObj);
                } else if (cellObj instanceof IPicture) {
                    picCells.add((IPicture) cellObj);
                } else if (cellObj instanceof IFormula) {
                    formulaCells.add((IFormula) cellObj);
                } else if (cellObj instanceof ICrossSum) {
                    crossSumCells.add((ICrossSum) cellObj);
                } else if (cellObj instanceof ISubTotal) {
                    subtotalCells.add((ISubTotal) cellObj);
                } else {
                    others.add(cellObj);
                }
            }
        }

        allCells.addAll(headerCells);
        allCells.addAll(footerCells);
        allCells.addAll(normalCells);
        allCells.addAll(subtotalCells);
        allCells.addAll(crossTabCells);
        allCells.addAll(crossSumCells);
        allCells.addAll(opSumCells);
        allCells.addAll(constCells);
        allCells.addAll(picCells);
        allCells.addAll(formulaCells);

        //校验模板信息
//        this.validate();

        this.scop();
        //补全信息
        this.completion();

        this.setLoaded(true);
    }

    /**
     * 补全模板信息
     */
    private void completion() {

        //普通报表补全信息
        if (this.sheetType == SheetType.NORMAL_SHEET) {
            completionNormal();
        }
    }

    /**
     * 补全普通报表模板信息
     */
    private void completionNormal() {
        //如果有小计，必须有且只有一个group标示的normal
        if (this.getAllSubtotals().size() > 0) {
            //标示group的normal个数
            int groupsize = 0;
            for (INormal iNormal : this.normalCells) {
                if (iNormal.hasGroup()) {
                    groupsize++;
                }
            }
            Assert.isTrue(groupsize == 1,sm.getString("poi.normal.groupsizeerror",groupsize));
            //如果模板定义了小计并且只有一个group标示，更新小计标示为true
            this.setSubTotal(true);
        }
    }

    public List<IConst> getAllConst() {
        return this.constCells;
    }

    public List<ICrossTab> getAllCrossTab() {
        return this.crossTabCells;
    }

    public List<ICrossSum> getAllCrossSum() {
        return this.crossSumCells;
    }

    public List<IFooter> getAllFooter() {
        return this.footerCells;
    }

    public List<IHeader> getAllHeader() {
        return this.headerCells;
    }

    public List<INormal> getAllNormal() {
        return this.normalCells;
    }

    public List<IOpSum> getAllOpSum() {
        return this.opSumCells;
    }

    public SheetType getSheetType() {
        return this.sheetType;
    }

    public void build(List<BuildStep> steps) throws POIException {
        for (BuildStep step : steps) {
            step.build(this);
        }
    }

    public boolean isLoaded() {
        return loaded;
    }

    protected void setLoaded(boolean loaded) {
        this.loaded = loaded;
    }

    public void scop() {
        if (!this.normalCells.isEmpty()) {
            this.sheetType = SheetType.NORMAL_SHEET;
        } else if (!this.crossTabCells.isEmpty()) {
            this.sheetType = SheetType.CROSSTAB_SHEET;
        } else if (allCells.size() == (this.headerCells.size()
                + this.footerCells.size() + this.constCells.size()
                + this.picCells.size() + this.formulaCells.size())) {
            this.sheetType = SheetType.CUST_1;
        } else {
            this.sheetType = SheetType.CUST_1;
        }
    }

    public Sheet getSheet() {
        return sheet;
    }

    public List<?> getData() {
        return this.data;
    }

    /**
     * 得到excel行对象，如果行不存在，创建一新行。
     */
    public Row getRow(int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    public Map<String, Object> getMetadata() {
        return this.metaData;
    }

    public void setDataOffRow(int dataOffRow) {
        this.dataOffRow = dataOffRow;
    }

    public int getDataOffRow() {
        return this.dataOffRow;
    }

    public void setDataOffColumn(int dataOffColumn) {
        this.dataOffColumn = dataOffColumn;
    }

    public int getDataOffColumn() {
        return this.dataOffColumn;
    }

    public void addMergedRegion(CellRangeAddress range) {
        sheet.addMergedRegion(range);
    }

    public void addMergedRegion(int firstRow, int lastRow, int firstCol,
                                int lastCol) {
        addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol,
                lastCol));
    }

    /**
     * 支持的数据类型有： 2013-10-25之前 Short Integer Long Double BigDecimal Date Calendar
     * String 2013-10-25新增 Boolean Float
     */
    public void setCellValue(int row, int column, ICellObject cellObject, Object value) {
        Cell cell = getCell(row, column);
        cell.setCellType(Cell.CELL_TYPE_BLANK);// 设置单元格类型为空
        cell.setCellType(cellObject.getCellType());// 设置单元格实际类型

        if (cellObject.getCellType() == Cell.CELL_TYPE_NUMERIC && value != null) {
            if (value instanceof Short) {
                cell.setCellValue((Short) value);
            } else if (value instanceof Integer) {
                cell.setCellValue((Integer) value);
            } else if (value instanceof Long) {
                cell.setCellValue((Long) value);
            } else if (value instanceof Float) {
                //如果数值类型的值位数大于 com.cfcc.deptone.excel.util.ExcelConsts.NUMERICAL_VALID_DIGIT
                //转化为文本写入单元格
                cell.setCellValue((Float) value);
            } else if (value instanceof Double) {

                cell.setCellValue((Double) value);
            } else if (value instanceof BigDecimal) {
                //如果数值类型的值位数大于 com.cfcc.deptone.excel.util.ExcelConsts.NUMERICAL_VALID_DIGIT
                //转化为文本写入单元格
                BigDecimal bigDecimal = (BigDecimal) value;
                String tmpValue = POIExcelUtil.formateBigDecimal(bigDecimal);
                if (tmpValue.replace(".", "").length() > ExcelConsts.NUMERICAL_VALID_DIGIT) {
                    cell.setCellStyle(cellObject.getCellStyle());
                    // 设置单元格类型为文本
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    String pattern = cellObject.getDataFormat().replace("_","");
                    cell.setCellValue(POIExcelUtil.formateBigDecimal(bigDecimal, pattern));
                } else {
                    cell.setCellValue(bigDecimal.doubleValue());
                }
            } else if (value instanceof Date) {
                cell.setCellValue((Date) value);
            } else if (value instanceof Calendar) {
                cell.setCellValue((Calendar) value);
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else { // 所有其它类型都当作RealichTextString处理
                RichTextString richString = creationHelper
                        .createRichTextString(String
                                .valueOf(value));
                cell.setCellValue(richString);
            }
        } else {
            RichTextString richString = creationHelper
                    .createRichTextString(value == null ? "" : String
                            .valueOf(value));
            cell.setCellValue(richString);
        }
    }


    /**
     * 支持的数据类型有：
     * 2013-10-25之前 Short Integer Long Double BigDecimal Date Calendar String
     * 2013-10-25新增 Boolean Float
     */
    public void setCellValueAndStyle(int row, int column, ICellObject cellObject, Object value) {
        Cell cell = getCell(row, column);
        cell.setCellType(Cell.CELL_TYPE_BLANK);// 设置单元格类型为空
        cell.setCellType(cellObject.getCellType());// 设置单元格实际类型

        CellStyle cellStyle = cellObject.getCellStyle();
        if (cellObject.getCellType() == Cell.CELL_TYPE_NUMERIC && value != null) {
            if (value instanceof Short) {
                cell.setCellValue((Short) value);
            } else if (value instanceof Integer) {
                cell.setCellValue((Integer) value);
            } else if (value instanceof Long) {
                cell.setCellValue((Long) value);
            } else if (value instanceof Float) {
                //如果数值类型的值位数大于 com.cfcc.deptone.excel.util.ExcelConsts.NUMERICAL_VALID_DIGIT
                //转化为文本写入单元格
                cell.setCellValue((Float) value);
            } else if (value instanceof Double) {

                cell.setCellValue((Double) value);
            } else if (value instanceof BigDecimal) {
                //如果数值类型的值位数大于 com.cfcc.deptone.excel.util.ExcelConsts.NUMERICAL_VALID_DIGIT
                //转化为文本写入单元格
                BigDecimal bigDecimal = (BigDecimal) value;
                String tmpValue = POIExcelUtil.formateBigDecimal(bigDecimal);
                if (tmpValue.replace(".", "").length() > ExcelConsts.NUMERICAL_VALID_DIGIT) {
                    cell.setCellStyle(cellObject.getCellStyle());
                    // 设置单元格类型为文本
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    String pattern = cellObject.getDataFormat().replace("_","");
                    pattern = pattern.replace("\"","");
                    cell.setCellValue(POIExcelUtil.formateBigDecimal(bigDecimal, pattern));
                    //创建一个文本样式
                    if((cellStyle = bigdecimald2string.get(cellObject.toString())) == null){
                        cellStyle = POIExcelUtil.createTextCellStyle(cellObject);
                        bigdecimald2string.put(cellObject.toString(),cellStyle);
                    }
                } else {
                    cell.setCellValue(bigDecimal.doubleValue());
                }
            } else if (value instanceof Date) {
                cell.setCellValue((Date) value);
            } else if (value instanceof Calendar) {
                cell.setCellValue((Calendar) value);
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else { // 所有其它类型都当作RealichTextString处理
                RichTextString richString = creationHelper
                        .createRichTextString(String
                                .valueOf(value));
                cell.setCellValue(richString);
            }
        } else {
            RichTextString richString = creationHelper
                    .createRichTextString(value == null ? "" : String
                            .valueOf(value));
            cell.setCellValue(richString);
        }

        cell.setCellStyle(cellStyle);
    }

    public int getNumMergedRegions() {
        return sheet.getNumMergedRegions();
    }

    public CellRangeAddress getMergedRegion(int index) {
        return sheet.getMergedRegion(index);
    }

    public void removeMergedRegion(int index) {
        sheet.removeMergedRegion(index);
    }

    public Cell getCell(int rowIndex, int columnIndex) {
        Row row = this.getRow(rowIndex);
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    public String getCellStringValue(int rowIndex, int columnIndex) {
        return getCell(rowIndex, columnIndex).getStringCellValue();
    }

    public String getStringCellValue(int row, int column) {
        return getCell(row, column).getStringCellValue();
    }

    public void setRegionCellStyle(CellStyle style, CellRangeAddress range) {
        for (int fr = range.getFirstRow(); fr <= range.getLastRow(); fr++) {
            for (int fc = range.getFirstColumn(); fc <= range.getLastColumn(); fc++) {
                this.getCell(fr, fc).setCellStyle(style);
            }
        }
    }

    public void setCellRangeNCellStyle(ICellObject cellObject) {
        // 设置合并单元格
        CellRangeAddress cellRange = cellObject.getRangeAddress();
        if (cellRange != null) {
            cellRange.setFirstRow(cellRange.getFirstRow()
                    + this.getDataOffRow());
            cellRange.setLastRow(cellRange.getLastRow() + this.getDataOffRow());
            this.addMergedRegion(cellRange);
            // 设置单元格格式
            this.setRegionCellStyle(cellObject.getCellStyle(), cellRange);
        } else {
            // 设置单元格格式
            cellObject.setCellStyle();
        }
    }

    public List<IPicture> getAllPicture() {
        return this.picCells;
    }

    public List<IFormula> getAllFormula() {
        return this.formulaCells;
    }

    public List<ISubTotal> getAllSubtotals() {
        return this.subtotalCells;
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    public CreationHelper getCreationHelper() {
        return this.creationHelper;
    }

    public Boolean hasSubTotal() {
        return subtotal;
    }

    public void setSubTotal(Boolean subtotal) {
        this.subtotal = subtotal;
    }


    public int getDataRightColumn(){
        int dataRightColumn = 0;
        if(SheetType.CROSSTAB_SHEET.equals(this.getSheetType())){
            List<ICrossTab>  iCrossTabs = POIExcelUtil.getAllDataCross(this.getAllCrossTab());
            ICrossTab iCrossTab = iCrossTabs.get(0);
            dataRightColumn = iCrossTab.getOriginalColumn() + this.dataOffColumn;
        }else if(SheetType.NORMAL_SHEET.equals(this.getSheetType())){
            dataRightColumn = this.dataOffColumn;
        }
        return dataRightColumn;
    }

    public int getDataLeftColumn(){
        int dataLeftColumn = 0;
        if(SheetType.CROSSTAB_SHEET.equals(this.getSheetType())){
            List<ICrossTab>  iCrossTabs = POIExcelUtil.getAllRowCross(this.getAllCrossTab());
            ICrossTab iCrossTab = iCrossTabs.get(0);
            dataLeftColumn = iCrossTab.getOriginalColumn();
        }else if(SheetType.NORMAL_SHEET.equals(this.getSheetType())){
            dataLeftColumn = this.getAllNormal().get(0).getOriginalColumn();
        }
        return dataLeftColumn;
    }

    public int getColumnWidth(int columnIndex) {
        return this.sheet.getColumnWidth(columnIndex);
    }

    public void setColumnWidth(int columnIndex, int width) {
        this.sheet.setColumnWidth(columnIndex,width);
    }
}

