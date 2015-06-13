package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.util.Assert;
import com.cfcc.deptone.excel.gen.ExcelBuilder;
import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.poi.POIFactory;
import com.cfcc.deptone.excel.util.ExcelConsts;
import com.cfcc.deptone.excel.util.StringManager;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 构建报表抽象类
 * <p/>
 * 2012-3-31
 *
 * @author why
 */
public abstract class AbstractExcelBuilder implements ExcelBuilder {


    private static final Log LOGGER = LogFactory.getLog(AbstractExcelBuilder.class);
    protected StringManager sm = StringManager.getManager("com.cfcc.deptone.excel.gen.inner");
    protected StringManager config = StringManager.getManager("com.cfcc.deptone.excel.gen.inner", "poi-config");

    protected Workbook workbook;
    protected Map<String, Object> metaDataMap[];
    protected List<?> rptDataList[];

    /**
     * sheet页个数
     */
    protected int numberOfSheets = 1;
    /**
     * sheet页隐藏级别，默认为Workbook.SHEET_STATE_HIDDEN
     */
    protected int sheetHiddenLevel = Workbook.SHEET_STATE_HIDDEN;

    public AbstractExcelBuilder() {
        String hiddenlevel = config.getString("poi.template.sheet.hiddenlevel");
        sheetHiddenLevel = Integer.parseInt(hiddenlevel);
    }

    public int build(File reportPath, Map<String, Object> metaDataMap, List<?> rptDataList) throws POIException {
        return build(reportPath, new Map[]{metaDataMap}, new List[]{rptDataList});
    }

    public int build(File reportPath, Map<String, Object> metaData[], List<?> rptData[]) throws POIException {
        FileInputStream fileInputStream = null;
        FileOutputStream fileOutputStream = null;
        this.metaDataMap = metaData;
        this.rptDataList = rptData;

        try {
            //
            fileInputStream = new FileInputStream(reportPath);
            buildWorkBookInfo(fileInputStream);

            //构建报表构建信息
            buildExcelBuilderInfoList();

            int activeIndex = -1;
            for (ExcelBuilderInfo excelBuilderInfo : this.excelBuilderList) {

                if (excelBuilderInfo.hasData()) {
                    ISheet sheet = POIFactory.createSheet(workbook, excelBuilderInfo.getSheet());

                    //加载sheet中定义的所有信息；并把sheet分类，由不同的sheet构建类进行构建
                    sheet.load(excelBuilderInfo.getMetaDataMap(), excelBuilderInfo.getRptDataList());
                    SheetBuilder builder = SheetBuilderFactory.getBuilder(sheet);
                    builder.write();
                    //第一个有数据的sheet页为活动sheet页
                    if (activeIndex < 0) activeIndex = excelBuilderInfo.getIndex();
                } else {
                    //隐藏无数据的sheet页
                    workbook.setSheetHidden(excelBuilderInfo.getIndex(), getSheetHiddenLevel());
                }
            }

            //活动sheet也
            workbook.setActiveSheet(activeIndex);

            // 写入文件
            fileOutputStream = new FileOutputStream(reportPath);
            workbook.write(fileOutputStream);

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new POIException(e);
        } finally {
            try {
                if (fileInputStream != null) {
                    fileInputStream.close();
                }
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                LOGGER.error(e.getMessage(), e);
                throw new POIException(e);
            }
        }
        return 0;
    }

    /**
     * 报表构建信息集合。包含当前模板所有sheet页的构建信息。
     * <br />
     * 如果数据中包含ExcelConsts.REPORT_TEMPLATE_SHEET_NAME，有可能一个sheet模板会写两遍数据。
     */
    List<ExcelBuilderInfo> excelBuilderList = new ArrayList<ExcelBuilderInfo>(2);

    /**
     * 构建报表构建信息
     */
    protected void buildExcelBuilderInfoList() {

        //按顺序遍历所有sheet页，整理ExcelBuilderInfo
        for (int index = 0; index < numberOfSheets; index++) {
            Sheet sheet = this.workbook.getSheetAt(index);
            //如果包含报表数据，并且数据中不包含sheet名称参数：ExcelConsts.REPORT_TEMPLATE_SHEET_NAME
            //创建包含数据的“报表构建信息”。
            if (containReportData(index) && !this.metaDataMap[index].containsKey(ExcelConsts.REPORT_TEMPLATE_SHEET_NAME)) {
                excelBuilderList.add(new ExcelBuilderInfo(index, sheet, this.metaDataMap[index], this.rptDataList[index]));
            }
            //如果不包含报表数据，创建只包含sheet索引的“报表构建信息”。
            else {
                excelBuilderList.add(new ExcelBuilderInfo(index));
            }
        }

        //遍历所有数据，整理ExcelBuilderInfo
        for (int i = 0; i < this.metaDataMap.length; i++) {
            Object sheetname = this.metaDataMap[i].get(ExcelConsts.REPORT_TEMPLATE_SHEET_NAME);
            //如果存在sheetname参数，按name查找sheet对象，创建构建信息。
            if (sheetname != null) {
                Sheet sheet = this.workbook.getSheet(String.valueOf(sheetname));
                Assert.notNull(sheet, sm.getString("poi.template.sheet.notfound", sheetname));
                int index = this.workbook.getSheetIndex(String.valueOf(sheetname));
                //如果包含相同index的对象，替换现有对象为新对象
                if (excelBuilderList.contains(new ExcelBuilderInfo(index))) {
                    int indexOf = excelBuilderList.indexOf(new ExcelBuilderInfo(index));
                    excelBuilderList.remove(indexOf);
                    excelBuilderList.add(indexOf, new ExcelBuilderInfo(index, sheet, this.metaDataMap[i], this.rptDataList[i]));
                } else {
                    excelBuilderList.add(new ExcelBuilderInfo(index, sheet, this.metaDataMap[i], this.rptDataList[i]));
                }
            }
        }
    }

    /**
     * 按照数据集合添加ExcelBuilderInfo
     *
     * @param i          数据索引
     * @param sheet      sheet对象
     * @param sheetIndex sheet索引
     */
    private void addExcelBuilderInfoFromDate(int i, Sheet sheet, int sheetIndex) {
        ExcelBuilderInfo excelBuilderInfo = new ExcelBuilderInfo(sheetIndex);
        if (excelBuilderList.contains(excelBuilderInfo)) {
            int indexOf = excelBuilderList.indexOf(excelBuilderInfo);
            excelBuilderInfo = excelBuilderList.get(indexOf);
            //如果sheet对象为空，在集合中删除此对象
            if(excelBuilderInfo.getSheet() == null){
                excelBuilderList.remove(indexOf);
            }
            excelBuilderInfo.setSheet(sheet);
            excelBuilderInfo.setMetaDataMap(this.metaDataMap[i]);
            excelBuilderInfo.setRptDataList(this.rptDataList[i]);
            excelBuilderList.add(indexOf, excelBuilderInfo);
        } else {
            excelBuilderList.add(new ExcelBuilderInfo(sheetIndex, sheet, this.metaDataMap[i], this.rptDataList[i]));
        }
    }

    /**
     * 是否包含报表数据
     *
     * @param index 所标
     * @return true包含，否则false。
     */
    protected boolean containReportData(int index) {
        return index < this.metaDataMap.length;
    }

    public int getSheetHiddenLevel() {
        return sheetHiddenLevel;
    }

    /**
     * 创建Excel工作簿信息
     *
     * @param fileInputStream Excel文件流
     * @throws IOException
     * @throws InvalidFormatException
     */
    protected void buildWorkBookInfo(FileInputStream fileInputStream) throws IOException, InvalidFormatException {
        workbook = WorkbookFactory.create(fileInputStream);
        numberOfSheets = workbook.getNumberOfSheets();
    }

}
