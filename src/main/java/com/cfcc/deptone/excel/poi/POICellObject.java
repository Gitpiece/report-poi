package com.cfcc.deptone.excel.poi;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ICellObject;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.model.support.InitBean;
import com.cfcc.deptone.excel.poi.operation.POIMarginOperation;
import com.cfcc.deptone.excel.poi.operation.POIOperation;
import com.cfcc.deptone.excel.poi.operation.POIOperationParser;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class
		POICellObject implements ICellObject ,InitBean{

	private ISheet sheet;
	private Cell cell;
	/**模板中定义的原始字符串*/
	private String stringCellValue;

	private CellStyle style;
	private CellRangeAddress rangeAddress;

	private int row;
	private int column;
	private short height;

	protected String propertyName;
	
	private String dataFormat;
	protected int originalRow;
	protected int originalColumn;
	private String originalCellValue;
	private int cellType;

	Collection<POIOperation> poiOperation;
	protected POICellObject(ISheet sheet, Cell cell) throws POIException {
		this(sheet,cell,true);
	}
	
	protected POICellObject(ISheet sheet, Cell cell, boolean emptyV) throws POIException {
		this.cell = cell;
		this.stringCellValue = this.getCell().getStringCellValue();
		this.sheet = sheet;
		this.style = cell.getCellStyle();
		this.dataFormat = this.style.getDataFormatString();
		this.originalRow = cell.getRowIndex();
		this.originalColumn = cell.getColumnIndex();
		this.setCoordinate(originalRow,originalColumn);
		if(Cell.CELL_TYPE_STRING == cell.getCellType()){
			this.originalCellValue = cell.getStringCellValue();
		}
		if (ExcelConsts.CELL_FORMATE_TEXT.equals(dataFormat)||ExcelConsts.CELL_FORMATE_GENERAL.equals(dataFormat)) {
			this.cellType = Cell.CELL_TYPE_STRING;
		} else {
			this.cellType = Cell.CELL_TYPE_NUMERIC;
		}
		
		//追加批注扩展参数
		if(cell.getCellComment() != null){
			appendXOP(cell.getCellComment());
			cell.removeCellComment();
		}
		
		if(emptyV){
			cell.setCellValue("");
		}

	}

//	/**
//	 * 补全信息，子类如有需要可实现此方法。
//	 */
//	protected void completion(){
//
//	}

	protected void appendXOP(Comment comment) throws POIException {
		POIOperationParser poiOperationParser = new POIOperationParser();
		poiOperation = poiOperationParser.parseCacheAnnotations(comment.getString().getString());
	}

	public Collection<POIOperation> getPoiOperation() {
		return poiOperation;
	}

	public Collection<POIOperation> getMarginPoiOperation(){
		List<POIOperation> list = new ArrayList<POIOperation>();
		if(getPoiOperation() != null)
		for (POIOperation poiOperation :getPoiOperation()) {
			if (POIMarginOperation.OPERATION.equals(poiOperation.getOperation())) {
				list.add(poiOperation);
			}
		}
		return list;
	}

	public String getDataFormat() {
		return this.dataFormat;
	}

	public short getHeight() {
		return height;
	}

	public void setHeight(short height) {
		this.height = height;
	}

	public CellStyle getCellStyle() {
		return this.style;
	}

	public int getRow() {
		return row;
	}

	public void setCoordinate(int row, int column) {
		this.row = row;
		this.column = column;
	}

	public int getColumn() {
		return column;
	}

	public int getOriginalRow() {
		return this.originalRow;
	}

	public int getOriginalColumn() {
		return this.originalColumn;
	}

	public boolean isCellMerge() {
		return rangeAddress != null;
	}

	public ISheet getSheet() {
		return sheet;
	}

	public Cell getCell() {
		return cell;
	}
	
	public String getCurrentCellStringValue(){
		return sheet.getCell(this.row, this.column).getStringCellValue();
	}

	public String getPropertyName() {
		return this.propertyName;
	}

	public List<IPlaceHolder> getAllPlaceHolder() {
		return POIExcelUtil.toPlaceHolders(this.getOriginalCellValue());
	}

	public String getOriginalCellValue() {
		return this.originalCellValue;
	}

	public void setValue(Object value) {
		sheet.setCellValue(this.getRow(), this.getColumn(),this, value);
	}
	
	public void setRangeAddress(CellRangeAddress rangeAddress) {
		this.rangeAddress = rangeAddress;
	}

	public CellRangeAddress getRangeAddress() {
		return rangeAddress;
	}

	public void setCellStyle() {
		sheet.getCell(this.row, this.column).setCellStyle(this.getCellStyle());
	}
	
	public int getCellType() {
		return cellType;
	}

	public void afterInit() {
		// nothing to do
	}

	public String toString(){
		return this.stringCellValue;
	}
}
