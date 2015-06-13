package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.IPicture;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * 构建 sheet 图片步骤
 * 
 * @author WangHuanyu
 */
public class BuildPicture implements BuildStep {
	private static final Log LOGGER = LogFactory.getLog(BuildPicture.class);
	public String getName() {
		return "BuildPicture";
	}

	public void build(ISheet sheet) throws POIException {
		List<IPicture> allConst = sheet.getAllPicture();
		for (IPicture iPic : allConst) {

			List<IPlaceHolder> holders = iPic.getAllPlaceHolder();

			String originalStr = iPic.getOriginalCellValue();
			Object picPath = null;
			for (IPlaceHolder hoder : holders) {
				picPath = POIExcelUtil.replaceMetadata(originalStr, hoder.toPHString(), sheet.getMetadata().get(hoder.toArray()[1]));
			}
			
			//合并单元格
			CellRangeAddress cellRange = iPic.getRangeAddress();
			int offColumn=1,offRow=1;
			if (cellRange != null) {
				offColumn += cellRange.getLastColumn()-cellRange.getFirstColumn();
				offRow += cellRange.getLastRow()-cellRange.getFirstRow();
			}
			
			File file = new File(picPath.toString());
			if(file.exists()){
				
				//增加图片
				//why 图片需要调整，支持ss usermodel
				Drawing patriarch = sheet.getSheet().createDrawingPatriarch();
				ClientAnchor anchor1 = sheet.getCreationHelper().createClientAnchor();//0, 0, 0, 0, (short) 0, 0, (short) 8, 20
				anchor1.setCol1(iPic.getOriginalColumn());
				anchor1.setRow1(iPic.getOriginalRow());
				anchor1.setCol2(iPic.getOriginalColumn()+offColumn);
				anchor1.setRow2(iPic.getOriginalRow()+offRow);
				patriarch.createPicture(anchor1, sheet.getWorkbook().addPicture(loadImage(file), Workbook.PICTURE_TYPE_JPEG));
				
			}
		}
	}

	public static byte[] loadImage(File file) throws POIException {
		FileInputStream input = null;
		try {
			input = new FileInputStream(file);
			ByteArrayOutputStream utput = new ByteArrayOutputStream();
			byte buf[] = new byte[1024];
			for (;;) {
				int noBytesRead = input.read(buf);
				if (noBytesRead == -1) {
					return utput.toByteArray();
				}
				utput.write(buf, 0, noBytesRead);
			}
		} catch (IOException e) {
			LOGGER.error(e.getMessage(),e);
			//FIXME WHY 需要详细错误说明
			throw new POIException(e);
		}finally{
			if(input != null){
				try {
					input.close();
				} catch (IOException e) {
					LOGGER.error(e.getMessage(),e);
				}
			}
		}
	}

	public void afterRow() throws POIException {
		// do nothing
	}

	public void beforeRow() throws POIException {
		// do nothing
	}

}
