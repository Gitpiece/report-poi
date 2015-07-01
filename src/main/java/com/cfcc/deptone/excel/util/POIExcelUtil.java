/**
 *
 */
package com.cfcc.deptone.excel.util;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ICellObject;
import com.cfcc.deptone.excel.model.ICrossTab;
import com.cfcc.deptone.excel.model.IOpSum;
import com.cfcc.deptone.excel.model.IPlaceHolder;
import com.cfcc.deptone.excel.poi.POIFactory;
import org.apache.commons.beanutils.DynaBean;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.beans.PropertyDescriptor;
import java.io.*;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 2012-3-23
 *
 * @author WangHuanyu
 */
public class POIExcelUtil {
    private static StringManager sm = StringManager.getManager("com.cfcc.deptone.excel.gen.inner");

    private POIExcelUtil() {
        //single
    }

    private static final Log LOGGER = LogFactory.getLog(POIExcelUtil.class);

    /**
     * 字符串截取#{normal.name}或#{cross.row.name}或row.name
     *
     * @param str      原字符串
     * @param dataType 数据类型
     * @return 截取后字符串
     */
    public static String subString(String str, String dataType) {
        if (str.endsWith("}")) {
            return str.substring(str.indexOf(dataType) + dataType.length() + 1, str.indexOf("}"));
        } else {
            return str.substring(str.indexOf(dataType) + dataType.length() + 1);
        }
    }


    /**
     * 得到数据对象object 的 propertyName属性值。
     * <br />
     * 如果 #object 是 {@link org.apache.commons.beanutils.DynaBean} ，通过接口方法获取属性值，否则，通过反射方式获取属性值。
     *
     * @param object       数据对象
     * @param propertyName 属性
     * @param methodMap    反射方法缓存map
     * @return 属性值
     */
    public static Object getPropertyValue(Object object, String propertyName, Map<String, Method> methodMap) {
        //如果是如果是apache dynaBean，通过DynaBean接口获取属性值
        if (object instanceof DynaBean) {
            try {
                return ((DynaBean) object).get(propertyName);
            } catch (IllegalArgumentException ie) {
                LOGGER.error(ie.getMessage(), ie);
                Assert.isTrue(false, sm.getString("poi.data.propertynotexist", propertyName));
            }
        }
        //通过反射方式获取属性值
        Method method;
        Object obj = null;
        try {
            if (methodMap.get(propertyName) == null) {
                PropertyDescriptor pd = new PropertyDescriptor(propertyName, object.getClass());
                method = pd.getReadMethod();
                methodMap.put(propertyName, method);
            } else {
                method = methodMap.get(propertyName);
            }
            obj = method.invoke(object);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            Assert.isTrue(false, sm.getString("poi.data.propertynotexist", propertyName));
        }
        return obj;
    }

    /**
     * 返回格式化的Excel求和公式
     *
     * @param firstRow    第一行
     * @param firstColumn 第一列
     * @param lastRow     最后一行
     * @param lastColumn  最后一列
     * @return 公式字符串
     */
    public static String formatAsSumString(int firstRow, int firstColumn, int lastRow, int lastColumn) {
        StringBuilder sb = new StringBuilder();
        sb.append("SUM(");

        CellReference cellRefFrom = new CellReference(firstRow, firstColumn);
        CellReference cellRefTo = new CellReference(lastRow, lastColumn);
        sb.append(cellRefFrom.formatAsString());
        // for a single-cell reference return A1 instead of A1:A1
        if (!cellRefFrom.equals(cellRefTo)) {
            sb.append(':');
            sb.append(cellRefTo.formatAsString());
        }
        sb.append(")");
        return sb.toString();
    }

    /**
     * 返回格式化的Excel求和公式
     *
     * @param firstRow    第一行
     * @param firstColumn 第一列
     * @param lastRow     最后一行
     * @param lastColumn  最后一列
     * @param hassubtotal 是否有小计，如果有小计求和减半。
     * @return 公式字符串
     */
    public static String formatAsSumStringHalf(int firstRow, int firstColumn, int lastRow, int lastColumn, Boolean hassubtotal) {
        StringBuilder sb = new StringBuilder(formatAsSumString(firstRow, firstColumn, lastRow, lastColumn));
        if (hassubtotal) {
            sb.append("/2");
        }
        return sb.toString();
    }

    public static String format4Member(final Map<String, IOpSum> opSumCache, final String members) {
        StringBuilder sb = new StringBuilder();
        sb.append("");

        char[] chars = members.toCharArray();
        StringBuilder memberName = new StringBuilder();
        boolean inOp = false;
        for (int i = 0; i < chars.length; i++) {
            char c = chars[i];
            switch (c) {
                case '[':
                case ']':
                    continue;
                case '(':
                case ')':
                case '+':
                case '-':
                case '*':
                case '/':
                case ',':
                    inOp = true;
                    break;
                default:
                    memberName.append(c);
            }

            if (inOp || i == chars.length - 2) {
                IOpSum opsum = opSumCache.get(memberName.toString());
                CellReference cellRef = new CellReference(opsum.getRow(), opsum.getColumn());

                sb.append(cellRef.formatAsString());

                // empty member name
                memberName.delete(0, memberName.length());
            }
            if (inOp && i < chars.length) {
                sb.append(c);
                inOp = false;
            }
        }
        sb.append("");
        return sb.toString();
    }

    /**
     * 将原标签用显示值替换
     *
     * @param str   标签 含{footer.name}
     * @param map   标题集合
     * @param refer 显示值
     * @return 替换后字符串
     */
    public static String changeMetadata(String str, Map<String, String> map, String refer) {
        String metadata = "#{" + refer + ".";
        String tmp = str.substring(str.indexOf(metadata) + metadata.length(), str.indexOf("}"));
        return str.replaceAll("\\#\\{" + refer + "." + tmp + "\\}", map.get(tmp));
    }

    public static Object replaceMetadata(String originalStr, String placeHolder, Object value) {
        if (originalStr.equalsIgnoreCase(placeHolder)) {
            return value;
        }
        return originalStr.replace(placeHolder, value == null ? "" : value.toString());
    }

    public static boolean isEmpty(String arg0) {
        return arg0 == null || "".equals(arg0.trim());
    }

    public static String toString(String arg0) {
        return arg0 == null || "null".equals(arg0) ? "" : arg0;
    }

    public static String[] toStrArray(String arg0) {
        return arg0.replaceAll("[#|{|}]", "").split("[#|{|}|\\.]");
    }

    public static List<IPlaceHolder> toPlaceHolders(final String str) {
        List<IPlaceHolder> placeHolderList = new ArrayList<IPlaceHolder>();

        boolean hasPH = true;
        String remainStr = str;
        while (hasPH) {

            int beginIndex = remainStr.indexOf("#{");
            int endIndex = remainStr.substring(beginIndex + 1).indexOf("}");
            if (beginIndex == -1 || endIndex == -1) {

                hasPH = false;
                continue;
            }

            placeHolderList.add(POIFactory.createPlaceHolder(remainStr.substring(beginIndex, endIndex + beginIndex + 2)));

            remainStr = remainStr.substring(endIndex + beginIndex + 2);
        }

        return placeHolderList;
    }

    public static String formateBigDecimal(BigDecimal bd,String pattern) {
        DecimalFormat df = new DecimalFormat(pattern);
        return df.format(bd);
    }

    public static String formateBigDecimal(BigDecimal bd) {
        return formateBigDecimal(bd, "###0.00");
    }

    /**
     * 保留cellObject 单元格的其他样式设置，修改为文本单元格样式，
     *
     * @param cellObject 单元格对象
     * @return 单元格样式
     */
    public static CellStyle createTextCellStyle(ICellObject cellObject) {
        CellStyle cs = cellObject.getSheet().getWorkbook().createCellStyle();
        cs.cloneStyleFrom(cellObject.getCellStyle());
        cs.setDataFormat(cellObject.getSheet().getWorkbook().createDataFormat().getFormat("@"));
        return cs;
    }

    /**
     * 创建空的单元格样式对象。
     *
     * @param workbook 工作簿
     * @return 单元格样式
     */
    public static CellStyle createNewCellStyle(Workbook workbook) {
        return workbook.createCellStyle();
    }

    /**
     * 创建相同的单元格样式对象。先从工作簿中查找是否有相同的样式对象，如果有返回现有的，如果没有穿件新的样式对象。
     * <p>
     * 注意：复制CellStyle时，2003支持的比较好，2007有部分样式复制不完整，如背景色，而且打开excel时有错误提示。
     * </p>
     * <p>
     * bug：如果复制的excel过大，可能会超过工作簿样式表的最大限制4000。
     * </p>
     *
     * @param workbook        工作簿
     * @param sourceWorkbook  原工作簿
     * @param sourceCellstyle 样式
     * @return 单元格样式
     */
    public static CellStyle createDuplicateCellStyle(Workbook workbook, Workbook sourceWorkbook, CellStyle sourceCellstyle) {
        short numCellstyle = workbook.getNumCellStyles();
        CellStyle cellStyle = null;
        while (numCellstyle-- > 0) {
            cellStyle = workbook.getCellStyleAt(numCellstyle);
            if (cellStyle.equals(sourceCellstyle)) {
                break;
            }

            cellStyle = null;
        }

        if (cellStyle == null) {
            cellStyle = createNewCellStyle(workbook);
            cellStyle.cloneStyleFrom(sourceCellstyle);
        }

//        //border
//        cellStyle.setAlignment(sourceCellstyle.getAlignment());
//        cellStyle.setBorderBottom(sourceCellstyle.getBorderBottom());
//        cellStyle.setBorderLeft(sourceCellstyle.getBorderLeft());
//        cellStyle.setBorderRight(sourceCellstyle.getBorderRight());
//        cellStyle.setBorderTop(sourceCellstyle.getBorderTop());
//        cellStyle.setBottomBorderColor(sourceCellstyle.getBottomBorderColor());
//        cellStyle.setBorderTop(sourceCellstyle.getBorderTop());
//        cellStyle.setBorderRight(sourceCellstyle.getBorderRight());
//        cellStyle.setBorderLeft(sourceCellstyle.getBorderLeft());
//        cellStyle.setBorderBottom(sourceCellstyle.getBorderBottom());
//
//        //font
//        Font font = workbook.createFont();
//        Font sourceFont = sourceWorkbook.getFontAt(sourceCellstyle.getFontIndex());
//        font.setBoldweight(sourceFont.getBoldweight());
//        font.setCharSet(sourceFont.getCharSet());
//        font.setColor(sourceFont.getColor());
//        font.setFontHeight(sourceFont.getFontHeight());
//        font.setFontHeightInPoints(sourceFont.getFontHeightInPoints());
//        font.setFontName(sourceFont.getFontName());
//        font.setItalic(sourceFont.getItalic());
//        font.setStrikeout(sourceFont.getStrikeout());
//        font.setTypeOffset(sourceFont.getTypeOffset());
//        font.setUnderline(sourceFont.getUnderline());
//        cellStyle.setFont(font);
////        DataFormat dataFormat = workbook.createDataFormat();
////        dataFormat.
////        cellStyle.setDataFormat();//sourceCellstyle.getDataFormatString();
//
//        //color
//        cellStyle.setTopBorderColor(sourceCellstyle.getTopBorderColor());
//        cellStyle.setBottomBorderColor(sourceCellstyle.getBottomBorderColor());
//        cellStyle.setLeftBorderColor(sourceCellstyle.getLeftBorderColor());
//        cellStyle.setRightBorderColor(sourceCellstyle.getRightBorderColor());
//        cellStyle.setFillBackgroundColor(sourceCellstyle.getFillBackgroundColor());
//        cellStyle.setFillForegroundColor(sourceCellstyle.getFillForegroundColor());
//
//        //others
//        cellStyle.setWrapText(sourceCellstyle.getWrapText());
//        cellStyle.setRotation(sourceCellstyle.getRotation());
//        cellStyle.setDataFormat(sourceCellstyle.getDataFormat());

        return cellStyle;
    }

    /**
     * 合并excel文件，合并非隐藏sheet页。
     *
     * @param toMerge   被合并的文件
     * @param mergeFile 合并文件
     * @return 0成功，其他错误。
     * @see #copyFile(File, File)
     * @see #deleteHiddenSheet(Workbook)
     * @see #copySheet(Workbook, Workbook, int)
     */
    public static int mergeExcel(String[] toMerge, String mergeFile) throws POIException {

        FileInputStream mergefileInputStream = null;
        FileOutputStream mergefileOutputStream = null;

        //复制 toMerge中的第一个文件，作为合并文件。
        copyFile(toMerge[0], mergeFile);
        //删除合并文件中的隐藏sheet页。
        Workbook mergeWorkbook = null;
        try {
            mergefileInputStream = new FileInputStream(mergeFile);
            mergeWorkbook = WorkbookFactory.create(mergefileInputStream);
            deleteHiddenSheet(mergeWorkbook);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } finally {
            if (mergefileInputStream != null) {
                try {
                    mergefileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        //从地二个toMerge开始复制非隐藏sheet
        for (int i = 1; i < toMerge.length; i++) {
            Workbook workbook = null;
            FileInputStream fileInputStream = null;
            try {
                fileInputStream = new FileInputStream(toMerge[i]);
                workbook = WorkbookFactory.create(fileInputStream);
                int numberOfSheets = workbook.getNumberOfSheets();
                for (int j = 0; j < numberOfSheets; j++) {
                    //复制非隐藏的sheet页
                    if (!workbook.isSheetHidden(j) && !workbook.isSheetVeryHidden(j)) {
                        copySheet(mergeWorkbook, workbook, j);
                    }
                }

                // 写入文件
                mergefileOutputStream = new FileOutputStream(mergeFile);
                mergeWorkbook.write(mergefileOutputStream);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            } finally {
                if (mergefileInputStream != null) {
                    try {
                        if (mergefileInputStream == null) {
                            mergefileInputStream.close();
                        }
                        if (mergefileOutputStream == null) {
                            mergefileOutputStream.close();
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }

        return 0;
    }

    /**
     * 复制sheet页内容
     *
     * @param mergeWorkbook 目标工作簿
     * @param workbook      原工作簿
     * @param index         原sheet索引
     */
    private static void copySheet(Workbook mergeWorkbook, Workbook workbook, int index) {
        Sheet sourcesheet = workbook.getSheetAt(index);
        Sheet targetSheet = mergeWorkbook.createSheet();
        int firstRowNum = sourcesheet.getFirstRowNum();
        int lashRowNum = sourcesheet.getLastRowNum();
        for (int rownum = firstRowNum; rownum <= lashRowNum; rownum++) {

            //设置行属性
            Row sourceRow = sourcesheet.getRow(rownum);
            if (sourceRow == null) {
                continue;
            }
            //复制行信息
            Row targetRow = targetSheet.createRow(rownum);
            targetRow.setHeight(sourceRow.getHeight());
            if (sourceRow.getRowStyle() != null)
                targetRow.setRowStyle(createDuplicateCellStyle(mergeWorkbook, workbook, sourceRow.getRowStyle()));
            targetRow.setZeroHeight(sourceRow.getZeroHeight());

            //设置列属性
            int lastcolumn = sourcesheet.getRow(rownum).getLastCellNum();
            for (int columnnum = sourcesheet.getRow(rownum).getFirstCellNum(); columnnum < lastcolumn; columnnum++) {

                //设置列宽
                targetSheet.setColumnWidth(columnnum, sourcesheet.getColumnWidth(columnnum));

                Cell sourceCell = sourcesheet.getRow(rownum).getCell(columnnum);
                if (sourceCell == null) {
                    continue;
                }

                //复制单元格信息
                Cell targetCell = targetRow.createCell(columnnum);
                targetCell.setCellComment(sourceCell.getCellComment());
                if (sourceCell.getCellStyle() != null) {
                    targetCell.setCellStyle(createDuplicateCellStyle(mergeWorkbook, workbook, sourceCell.getCellStyle()));
                }
                if (sourceCell.getHyperlink() != null) {
                    targetCell.setHyperlink(sourceCell.getHyperlink());
                }
                targetCell.setCellType(sourceCell.getCellType());
                if (Cell.CELL_TYPE_STRING == sourceCell.getCellType()) {
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                } else if (Cell.CELL_TYPE_NUMERIC == sourceCell.getCellType()) {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                } else if (Cell.CELL_TYPE_FORMULA == sourceCell.getCellType()) {
                    targetCell.setCellFormula(sourceCell.getCellFormula());
                } else if (Cell.CELL_TYPE_BLANK == sourceCell.getCellType()) {
                    //do nothing
                } else if (Cell.CELL_TYPE_BOOLEAN == sourceCell.getCellType()) {
                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                } else if (Cell.CELL_TYPE_ERROR == sourceCell.getCellType()) {
                    // do nothing
                }
            }
        }

        //设置合并单元格
        int mergeNum = sourcesheet.getNumMergedRegions();
        for (int i = 0; i < mergeNum; i++) {
            targetSheet.addMergedRegion(sourcesheet.getMergedRegion(i));
        }
    }

    /**
     * 删除隐藏的sheet页
     *
     * @param mergeWorkbook excel工作簿
     */
    public static void deleteHiddenSheet(Workbook mergeWorkbook) {
        try {
            int numberOfSheets = mergeWorkbook.getNumberOfSheets();
            int i = numberOfSheets;
            while (--i >= 0) {
                if (mergeWorkbook.isSheetHidden(i)) {
                    mergeWorkbook.removeSheetAt(i);
                }
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }

    }

    /**
     * 复制文件
     *
     * @param sourceFile 原文件
     * @param targetFile 目标文件
     * @return 0成功，其他失败
     */
    public static int copyFile(File sourceFile, File targetFile) throws POIException {
        byte[] byteArray = new byte[1024];
        FileInputStream fileInputStream = null;
        FileOutputStream fileOutputStream = null;
        try {
            fileInputStream = new FileInputStream(sourceFile);
            fileOutputStream = new FileOutputStream(targetFile);

            int length;
            while ((length = fileInputStream.read(byteArray)) > 0) {
                fileOutputStream.write(byteArray, 0, length);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (fileInputStream != null) {
                    fileInputStream.close();
                }
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                throw new POIException(e);
            }
        }
        return 0;
    }

    /**
     * 复制文件
     */
    public static int copyFile(String source, String target) throws POIException {
        return copyFile(new File(source), new File(target));
    }

    //****************************************************
    // 交叉报表工具方法
    //****************************************************
    //

    public static List<ICrossTab> getAllDataCross(final List<ICrossTab> allCrossTab) {
        List<ICrossTab> allRowCrossTab = new ArrayList<ICrossTab>();
        for (ICrossTab crossTab : allCrossTab) {
            if (crossTab.isData()) {
                allRowCrossTab.add(crossTab);
            }
        }
        return allRowCrossTab;
    }


    public static List<ICrossTab> getAllRowCross(final List<ICrossTab> allCrossTab) {
        List<ICrossTab> allRowCrossTab = new ArrayList<ICrossTab>();
        for (ICrossTab crossTab : allCrossTab) {
            if (crossTab.isRow()) {
                allRowCrossTab.add(crossTab);
            }
        }
        return allRowCrossTab;
    }


    public static List<ICrossTab> getAllColumnCross(final List<ICrossTab> allCrossTab) {
        List<ICrossTab> allRowCrossTab = new ArrayList<ICrossTab>();
        for (ICrossTab crossTab : allCrossTab) {
            if (crossTab.isColumn()) {
                allRowCrossTab.add(crossTab);
            }
        }
        return allRowCrossTab;
    }

    /**
     * 获取excel文件的所有sheet页名称
     *
     * @param filepath excel文件绝对路径
     * @return excel所有sheet页名称数组
     */
    public static String[] getAllSheetName(String filepath) throws POIException {

        FileInputStream fileInputStream = null;
        try {
            //
            fileInputStream = new FileInputStream(filepath);
            Workbook workbook = WorkbookFactory.create(fileInputStream);
            int numberofsheet = workbook.getNumberOfSheets();
            String[] sheetnames = new String[numberofsheet];
            for (int index = 0; index < numberofsheet; index++) {
                sheetnames[index] = workbook.getSheetName(index);
            }
            return sheetnames;
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new POIException(e);
        } finally {
            try {
                if(fileInputStream != null){
                    fileInputStream.close();
                }
            } catch (IOException e) {
                LOGGER.error(e.getMessage(), e);
            }
        }
    }

}
