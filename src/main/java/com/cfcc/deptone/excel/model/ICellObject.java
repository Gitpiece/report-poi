package com.cfcc.deptone.excel.model;

import com.cfcc.deptone.excel.poi.operation.POIOperation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collection;
import java.util.List;

/**
 * 单元格，这个接口是一个基本接口，头、尾等类型接口需要继承此接口。
 *
 * @author WangHuanyu
 */
public interface ICellObject {

    /**
     * 属性名
     *
     * @return property name
     */
    String getPropertyName();

    /**
     * 是否是合并单元格
     */
    boolean isCellMerge();

    /**
     * 单元格范围坐标
     *
     * @param rangeAddress cell range address object.
     */
    void setRangeAddress(CellRangeAddress rangeAddress);

    /**
     * 获取单元格范围坐标
     *
     * @return current cell range address
     */
    CellRangeAddress getRangeAddress();

    /**
     * 得到当前行
     *
     * @return current row
     */
    int getRow();

    /**
     * 设置当前坐标
     *
     * @param row    行
     * @param column 列
     */
    void setCoordinate(int row, int column);

    /**
     * 设置当前单元格值
     *
     * @param value 值
     */
    void setValue(Object value);

    /**
     * 得到当前列
     *
     * @return 当前单元格column
     */
    int getColumn();

    /**
     * 得到原始行
     *
     * @return 返回原始行
     */
    int getOriginalRow();

    /**
     * 得到原始列
     *
     * @return 返回原始列
     */
    int getOriginalColumn();

    /**
     * 得到原始值
     *
     * @return 返回模板中单元格中的字符
     */
    String getOriginalCellValue();

    /**
     * 得到数据格式
     *
     * @return 得到单元格格式定义字符串
     */
    String getDataFormat();

    /**
     * 得到单元格样式
     *
     * @return 返回poi单元格样式对象
     */
    CellStyle getCellStyle();

    /**
     * 得到poi单元格对象
     *
     * @return 返回poi单元格对象
     */
    Cell getCell();

    /**
     * 得到当前单元格的内容，当前单元格的坐标是通过 {@link #getRow()}和{@link #getColumn()}方法得到。
     *
     * @return 返回当前单元格内容
     */
    String getCurrentCellStringValue();

    /**
     * 设置单元格样式，坐标是通过 {@link #getRow()}和{@link #getColumn()}方法得到，样式通过{@link #getCellStyle()}方法得到。
     */
    void setCellStyle();

    /**
     * 得到当前单元格的占位符
     *
     * @return 当前单元格占位符集合
     */
    List<IPlaceHolder> getAllPlaceHolder();

    /**
     * 得到单元格类型
     *
     * @return 单元格类型
     * @see org.apache.poi.ss.usermodel.Cell
     */
    int getCellType();

    /**
     * 获取工作表对象
     *
     * @return 工作簿接口对象
     */
    ISheet getSheet();

    /**
     * Poi扩展操作参数
     *
     * @return Poi扩展操作参数
     */
    Collection<POIOperation> getPoiOperation();

    /**
     * Poi margin扩展操作参数
     *
     * @return Poi扩展操作参数
     */
    Collection<POIOperation> getMarginPoiOperation();
    /**
     //	Color getColor();
     //	Font getFont();
     //	String getSemanticName();
     //	String getRegion();
     //	String getQualifyN();
     //	String getCellValue();
     //	void setValue(int row, int column, Object value);
     */
}
