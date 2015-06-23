package com.cfcc.deptone.excel.model;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.gen.inner.BuildStep;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.Map;

/**
 * Excel Sheet接口
 *
 * @author WangHuanyu
 */
public interface ISheet {
    /*============================
     * 得到单元格数据集合
     ============================*/
    List<IConst> getAllConst();

    List<ICrossTab> getAllCrossTab();

    List<ICrossSum> getAllCrossSum();

    List<IFooter> getAllFooter();

    List<IHeader> getAllHeader();

    List<INormal> getAllNormal();

    List<ISubTotal> getAllSubtotals();

    List<IOpSum> getAllOpSum();

    List<IPicture> getAllPicture();

    List<IFormula> getAllFormula();

    /**
     * 得到sheet type
     *
     * @return sheet type
     */
    SheetType getSheetType();

    /**
     * 加载所有模板定义信息。
     *
     * @param metaData meta data
     * @param data     数据集
     * @throws POIException
     */
    void load(Map<String, Object> metaData, List<?> data) throws POIException;

    /**
     * 根据加载的模板信息判断，确定sheet的类型。
     */
    void scop();

    /**
     * 按定义好的步骤写入数据。
     *
     * @param steps 构建步骤集合
     */
    void build(List<BuildStep> steps) throws POIException;

    /**
     * 是否加载完成
     *
     * @return 是否加载完成
     * @throws Exception 异常
     */
    boolean isLoaded() throws POIException;

    /**
     * 得到要写入的数据集合
     *
     * @return 数据
     */
    List<?> getData();

    /**
     * 得到元数据集合
     *
     * @return meta data
     */
    Map<String, Object> getMetadata();

    /**
     * 得到excel行对象
     *
     * @param i 行
     * @return POI行对象
     */
    Row getRow(int i);

    /**
     * 得到excel单元格对象
     *
     * @param row    行
     * @param column 列
     * @return POI单元格对象
     */
    Cell getCell(int row, int column);

    /**
     * 得到单元格字符串，与{@link #getStringCellValue(int, int)}重复。
     *
     * @param row    行
     * @param column 列
     * @return 单元格内容
     */
    String getCellStringValue(int row, int column);

    /**
     * 得到单元格内容，与{@link #getCellStringValue(int, int)}重复。
     *
     * @param row    行
     * @param column 列
     * @return 单元格内容
     */
    String getStringCellValue(int row, int column);

    /**
     * 设置写入数据行数
     *
     * @param dataOffRow 数据写入行数
     */
    void setDataOffRow(int dataOffRow);

    /**
     * 得到写入数据行数
     *
     * @return 数据写入行数
     */
    int getDataOffRow();

    /**
     * 设置写入数据列数
     *
     * @param dataOffColumn 数据写入列数
     */
    void setDataOffColumn(int dataOffColumn);

    /**
     * 得到写入数据列数
     *
     * @return 数据写入列数
     */
    int getDataOffColumn();

    /**
     * 增加合并区域
     *
     * @param range cell range address
     */
    void addMergedRegion(CellRangeAddress range);

    /**
     * 增加合并区域
     *
     * @param firstRow 第一行
     * @param lastRow  最后一行
     * @param firstCol 第一列
     * @param lastCol  最后一列
     */
    void addMergedRegion(int firstRow, int lastRow, int firstCol, int lastCol);

    /**
     * 得到POI sheet对象
     *
     * @return poi sheet对象
     */
    Sheet getSheet();

    /**
     * 得到工作簿 对象
     *
     * @return 工作簿
     */
    Workbook getWorkbook();

    /**
     * 设置单元格值
     *
     * @param row        行
     * @param column     列
     * @param cellObject 单元格对象
     * @param value      值
     */
    void setCellValue(int row, int column, ICellObject cellObject, Object value);

    /**
     * 得到sheet中region个数
     *
     * @return sheet中region的个数
     */
    int getNumMergedRegions();

    /**
     * 得到sheet中合并单元格
     *
     * @param index sheet中的region索引
     * @return cell range address
     */
    CellRangeAddress getMergedRegion(int index);

    /**
     * 移除sheet中合并单元格
     *
     * @param index sheet中的region索引
     */
    void removeMergedRegion(int index);

    /**
     * 设置区域内单元格的样式
     *
     * @param style 单元格样式
     * @param range 单元格range
     */
    void setRegionCellStyle(CellStyle style, CellRangeAddress range);

    /**
     * 设置合并单元格和样式
     *
     * @param cellObject 单元格对象
     */
    void setCellRangeNCellStyle(ICellObject cellObject);

    /**
     * @return 返回CreationHelper
     */
    CreationHelper getCreationHelper();

    /**
     * 报表是否有小计。目前小计只支持normal报表。
     * @return
     */
    Boolean hasSubTotal();

    /**
     * 得到真实的数据最右边的列
     * @return
     */
    int getDataRightColumn();

    /**
     * 得到真实的数据最右边的列
     * @return
     */
    int getDataLeftColumn();

    /**
     * 获取列宽
     * @param columnIndex - the column to set (0-based)
     * @return width - the width in units of 1/256th of a character width
     */
    int getColumnWidth(int columnIndex);
    /**
     * 设置列宽
     * @param columnIndex - the column to set (0-based)
     * @param width - the width in units of 1/256th of a character width
     * @throws IllegalArgumentException if width > 255*256 (the maximum column width in Excel is 255 characters)
     */
    void setColumnWidth(int columnIndex, int width);
}
