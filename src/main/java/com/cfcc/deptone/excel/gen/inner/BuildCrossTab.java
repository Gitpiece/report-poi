package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ICrossTab;
import com.cfcc.deptone.excel.model.ISheet;
import com.cfcc.deptone.excel.poi.POICrossTab;
import com.cfcc.deptone.excel.poi.operation.POIOperation;
import com.cfcc.deptone.excel.poi.operation.POIOrderOperation;
import com.cfcc.deptone.excel.poi.support.CrossTabData;
import com.cfcc.deptone.excel.poi.support.CrossTabHeader;
import com.cfcc.deptone.excel.poi.support.CrossTabTreeNode;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import com.cfcc.deptone.excel.util.TreeNode;

import java.lang.reflect.Method;
import java.util.*;

/**
 * 构建交叉报表。
 * 2015-3-5 完全修改交叉报表实现方式
 *
 * @author wanghuanyu
 */
public class BuildCrossTab implements BuildStep {

    private ISheet sheet;
    private Map<String, Method> methodMap = new HashMap<String, Method>();

    private List<ICrossTab> ccolomnhead;
    private List<ICrossTab> crowhead;
    private List<ICrossTab> cdata;//目前只支持一个数据单元格
    /**
     * 交叉报表的列表头
     */
    private CrossTabTreeNode<CrossTabHeader> columntree;
    /**
     * 交叉报表的行表头
     */
    private CrossTabTreeNode<CrossTabHeader> rowtree;
    /**
     * 数据矩阵，用于保存数据
     */
    private CrossTabData[][] dataMatrix;

    public String getName() {
        return "BuildCrossTab";
    }

    public void build(ISheet sheet) throws POIException {
        this.sheet = sheet;
        ccolomnhead = POIExcelUtil.getAllColumnCross(sheet.getAllCrossTab());
        crowhead = POIExcelUtil.getAllRowCross(sheet.getAllCrossTab());
        cdata = POIExcelUtil.getAllDataCross(sheet.getAllCrossTab());
        //数据列宽
        int datacolumnwidth = this.sheet.getColumnWidth(cdata.get(0).getOriginalColumn());

        columntree = this.buildCrossTabHeader(ccolomnhead, sheet.getData());
        rowtree = this.buildCrossTabHeader(crowhead, sheet.getData());

        //计算出报表的行列数
        int rowNum = leafCount(rowtree);
        int columnNum = leafCount(columntree);
        //set off row/column
        sheet.setDataOffRow(rowNum - 1);
        sheet.setDataOffColumn(columnNum - 1);

        //创建数据矩阵，用于保存数据
        dataMatrix = new CrossTabData[rowNum][columnNum];

        //把数据填入数据矩阵
        fillDataMatrix(dataMatrix, sheet.getData());

        //把数据矩阵写入excel
        writeDataMatrix(dataMatrix,datacolumnwidth);
        //把表头写入excel
        writeColumnHead(0, columntree);
        writeRowHead(0, rowtree);
        this.mergedRegion(sheet, sheet.getDataOffRow(), sheet.getDataOffColumn());
    }

    /**
     * @param x      偏移X
     * @param parent 父节点
     */
    private void writeColumnHead(int x, CrossTabTreeNode<CrossTabHeader> parent) {
        TreeNode<CrossTabHeader>[] treenode = parent.getAllChild();
        int offX = x, leafcount;
        for (TreeNode<CrossTabHeader> aTreenode : treenode) {
            CrossTabTreeNode<CrossTabHeader> treeNode2 = (CrossTabTreeNode<CrossTabHeader>) aTreenode;

            ICrossTab crossTab = treeNode2.getCrossTab();
            //得到当前节点的子节点个数
            leafcount = this.leafCount(treeNode2);
            int currentnodechildoffx = leafcount > 0 ? leafcount - 1 : leafcount;
            crossTab.setCoordinate(crossTab.getOriginalRow(), crossTab.getOriginalColumn() + offX
                    + currentnodechildoffx);
            crossTab.setCellStyle();
            sheet.setCellValue(crossTab.getRow(), crossTab.getColumn(), crossTab, treeNode2.getLabel());

            if (treeNode2.hasChild()) {
                writeColumnHead(offX, treeNode2);
            }
            offX += leafcount > 0 ? leafcount : 1;
        }
    }

    private void writeRowHead(int y, CrossTabTreeNode<CrossTabHeader> parent) {
        TreeNode<CrossTabHeader>[] treenode = parent.getAllChild();
        int offy = y, leafcount;
        for (TreeNode<CrossTabHeader> aTreenode : treenode) {
            CrossTabTreeNode<CrossTabHeader> treeNode2 = (CrossTabTreeNode<CrossTabHeader>) aTreenode;

            ICrossTab crossTab = treeNode2.getCrossTab();
            //得到当前节点的子节点个数
            leafcount = this.leafCount(treeNode2);
            int currentnodechildoffy = leafcount > 0 ? leafcount - 1 : leafcount;
            crossTab.setCoordinate(crossTab.getOriginalRow() + offy + currentnodechildoffy,
                    crossTab.getOriginalColumn());
            crossTab.setCellStyle();
            sheet.setCellValue(crossTab.getRow(), crossTab.getColumn(), crossTab, treeNode2.getLabel());

            if (treeNode2.hasChild()) {
                writeRowHead(offy, treeNode2);
            }
            offy += leafcount > 0 ? leafcount : 1;
        }
    }

    /**
     * 数据矩阵写入excel
     *
     * @param dataMatrix 数据矩阵
     * @param datacolumnwidth 数据列宽
     */
    private void writeDataMatrix(CrossTabData[][] dataMatrix,int datacolumnwidth) {
        boolean setdatacolumnwidth=false;
        for (int irow = 0; irow < dataMatrix.length; irow++) {
            CrossTabData[] rowDatas = dataMatrix[irow];

            for (int icolumn = 0; icolumn < rowDatas.length; icolumn++) {
                CrossTabData cellData = rowDatas[icolumn];
                //
                int row = cellData.getCrossTabDataPOI().getOriginalRow() + irow;
                int column = cellData.getCrossTabDataPOI().getOriginalColumn() + icolumn;
                Object value = cellData.getOriginalObj() != null ? POIExcelUtil.getPropertyValue(cellData.getOriginalObj(), cellData.getCrossTabDataPOI()
                        .getPropertyName(), methodMap) : "";
                cellData.getCrossTabDataPOI().setCoordinate(row, column);
                cellData.getCrossTabDataPOI().setCellStyle();
                cellData.getCrossTabDataPOI().setValue(value);
                //sheet.setCellValue(row, column, cellData.getCrossTabDataPOI(), value);

                //设置数据列宽
                if(!setdatacolumnwidth){
                    this.sheet.setColumnWidth(cellData.getCrossTabDataPOI().getOriginalColumn() + icolumn,datacolumnwidth);
                }
            }

            setdatacolumnwidth = true;
        }
    }

    /**
     * 合并交叉表行和列
     *
     * @param sheet     sheet对象
     * @param offrow    移动最大行
     * @param offcolumn 移动最大列
     */
    private void mergedRegion(ISheet sheet, int offrow, int offcolumn) {

        // 标题列合并
        for (ICrossTab crossTab : ccolomnhead) {
            int originalrow = crossTab.getOriginalRow(),
                    originalcolumn = crossTab.getOriginalColumn();
            int preColumnIndex = -1;//前一个不为空的坐标
            //从右往左处理合并
            for (int j = 0; j <= offcolumn; j++) {
                String cellValue = sheet.getStringCellValue(originalrow, originalcolumn + j);
                if (POIExcelUtil.isEmpty(cellValue) && preColumnIndex == -1) {
                    preColumnIndex = j;
                    continue;
                } else if (!POIExcelUtil.isEmpty(cellValue)) {
                    //无合并单元格
                    if (preColumnIndex == -1) {
                        continue;
                    }

                    sheet.setCellValue(originalrow, originalcolumn + j, crossTab, "");
                    crossTab.setCoordinate(originalrow, originalcolumn + preColumnIndex);
                    crossTab.setCellStyle();
                    sheet.setCellValue(originalrow, originalcolumn + preColumnIndex, crossTab, cellValue);
                    sheet.addMergedRegion(originalrow, originalrow, originalcolumn + preColumnIndex, originalcolumn + j);
                    //reset column index
                    preColumnIndex = -1;
                } else { //设置合并单元格的样式
                    sheet.setCellValue(originalrow, originalcolumn + j, crossTab, "");
                    crossTab.setCoordinate(originalrow, originalcolumn + j);
                    crossTab.setCellStyle();
                }
            }
        }

        // 标题行合并
        for (ICrossTab crossTab : crowhead) {
            int originalrow = crossTab.getOriginalRow(),
                    originalcolumn = crossTab.getOriginalColumn();
            int preRowIndex = -1;//前一个不为空的坐标
            //从右往左处理合并
            for (int j = 0; j <= offrow; j++) {
                String cellValue = sheet.getStringCellValue(originalrow + j, originalcolumn);
                if (POIExcelUtil.isEmpty(cellValue) && preRowIndex == -1) {
                    preRowIndex = j;
                    continue;
                } else if (!POIExcelUtil.isEmpty(cellValue)) {
                    //无合并单元格
                    if (preRowIndex == -1) {
                        continue;
                    }

                    sheet.setCellValue(originalrow + j, originalcolumn, crossTab, "");
                    crossTab.setCoordinate(originalrow + preRowIndex, originalcolumn);
                    crossTab.setCellStyle();
                    sheet.setCellValue(originalrow + preRowIndex, originalcolumn, crossTab, cellValue);
                    sheet.addMergedRegion(originalrow + preRowIndex, originalrow + j, originalcolumn, originalcolumn);
                    //reset column index
                    preRowIndex = -1;
                } else { //设置合并单元格的样式
                    sheet.setCellValue(originalrow + j, originalcolumn, crossTab, "");
                    crossTab.setCoordinate(originalrow + j, originalcolumn);
                    crossTab.setCellStyle();
                }
            }
        }
    }

    /**
     * 把数据填入数据矩阵。
     * 通过构建好的表头定位数据坐标，构建数据矩阵。
     *
     * @param dataMatrix 数据矩阵
     * @param data       报表数据集合
     */
    private void fillDataMatrix(CrossTabData[][] dataMatrix, List<?> data) {
        for (Object object : data) {
            CrossTabData crossTabData = new CrossTabData();
            crossTabData.setOriginalObj(object);
            crossTabData.setCrossTabDataPOI(this.cdata.get(0));//设置数据对应的交叉报表数据定义对象
            computePoint(crossTabData);
            dataMatrix[crossTabData.getPoint().getX()][crossTabData.getPoint().getY()] = crossTabData;
        }

        //fill empty element
        for (int x = 0; x < dataMatrix.length; x++) {
            CrossTabData[] rowArray = dataMatrix[x];
            for (int y = 0; y < rowArray.length; y++) {
                CrossTabData crossTabData = rowArray[y];
                if (crossTabData == null) {
                    CrossTabData cTabData = new CrossTabData();
                    cTabData.setCrossTabDataPOI(this.cdata.get(0));//设置数据对应的交叉报表数据定义对象
                    cTabData.setPoint(x, y);
                    dataMatrix[x][y] = cTabData;
                }
            }
        }
    }

    /**
     * 计算数据坐标
     *
     * @param crossTabData 数据对象
     */
    private void computePoint(CrossTabData crossTabData) {
        StringBuilder columnStr = new StringBuilder();
        for (ICrossTab aCcolomnhead : ccolomnhead) {
            POICrossTab crossTab = (POICrossTab) aCcolomnhead;
            columnStr.append(String.valueOf(POIExcelUtil.getPropertyValue(crossTabData.getOriginalObj(),
                    crossTab.getPropertyName(), methodMap)));
            columnStr.append(">");
        }
        StringBuilder rowStr = new StringBuilder();
        for (ICrossTab aCrowhead : crowhead) {
            POICrossTab crossTab = (POICrossTab) aCrowhead;
            rowStr.append(String.valueOf(POIExcelUtil.getPropertyValue(crossTabData.getOriginalObj(),
                    crossTab.getPropertyName(), methodMap)));
            rowStr.append(">");
        }
        //rowtree决定x坐标，columntree决定y坐标。
        crossTabData.setPoint(computeTreeCoordinate(rowtree, rowStr), computeTreeCoordinate(columntree, columnStr));
        crossTabData.setRowString(rowStr.toString());
        crossTabData.setColumnString(columnStr.toString());
    }

    /**
     * 计算树坐标
     *
     * @return 坐标值,left leaf count.
     */
    private int computeTreeCoordinate(TreeNode<CrossTabHeader> headroot, StringBuilder pathStr) {
        String[] pathArr = pathStr.toString().split(">");
        TreeNode<CrossTabHeader>[] child = headroot.getAllChild();
        int i = 0, leftLeafCount = 0, pathArrIndex = 0;
        while (true) {
            TreeNode<CrossTabHeader> treeNode = child[i++];
            //如果与路径元素匹配，把child更新为匹配节点的子节点，重置索引属性。
            if (treeNode.getLabel().equals(pathArr[pathArrIndex])) {
                child = treeNode.getAllChild();
                if (child.length == 0) {
                    break;
                }
                i = 0;//坐标清零
                pathArrIndex++;
            } else if (!child[0].hasChild()) { //如果节点是叶子节点，左叶子加1
                leftLeafCount++;
            } else { //否则，获取当前节点的叶子节点个数
                leftLeafCount += this.leafCount(treeNode);
            }
        }

        return leftLeafCount;
    }

    /**
     * 计算节点叶子节点个数
     *
     * @param root 根节点
     * @return 叶子个数
     */
    private int leafCount(TreeNode<CrossTabHeader> root) {
        TreeNode<CrossTabHeader>[] childs = root.getAllChild();
        int count = 0;
        for (TreeNode<CrossTabHeader> treeNode : childs) {
            if (treeNode != null && treeNode.getAllChild().length > 0) {
                count += leafCount(treeNode);
            } else {
                count += childs.length;
                break;
            }
        }

        return count;
    }


    /**
     * 整理表头信息
     *
     * @param head 列表头集合
     * @param listdata 数据集合
     */
    private CrossTabTreeNode<CrossTabHeader> buildCrossTabHeader(List<ICrossTab> head, List listdata) {
        CrossTabTreeNode<CrossTabHeader> columnTreeRoot = new CrossTabTreeNode<CrossTabHeader>();

        //循环数据
        for (Object dataObj : listdata) {
            //每一个数据对象，通过配置的头信息创建一个树的分支
            CrossTabTreeNode<CrossTabHeader> currentTreeParent = columnTreeRoot;
            for (ICrossTab crossTab : head) {
                Object value = POIExcelUtil.getPropertyValue(dataObj, crossTab.getPropertyName(), methodMap);
                //如果当前节点不包含 label 是 value的节点，创建子节点，并设置currentTreeParent为currentTreeNode
                if (!currentTreeParent.containNode(String.valueOf(value))) {
                    currentTreeParent = new CrossTabTreeNode<CrossTabHeader>(currentTreeParent, String.valueOf(value),
                            new CrossTabHeader(crossTab
                                    .getPropertyName(), String
                                    .valueOf(value)));
                    currentTreeParent.setCrossTab(crossTab);
                } else { //否则，设置 currentTreeParent 为 label 是 value的子节点
                    currentTreeParent = (CrossTabTreeNode) currentTreeParent.getChild(String.valueOf(value));
                }
            }
        }
        doExtendOperation(head, columnTreeRoot);
        return columnTreeRoot;
    }

    //==================================================================
    //  交叉报表表头扩展操作
    //==================================================================

    /**
     * @param head           交叉报表表头
     * @param columnTreeRoot 交叉报表表头数据
     */
    private void doExtendOperation(List<ICrossTab> head, CrossTabTreeNode<CrossTabHeader> columnTreeRoot) {
        for (int i = 0; i < head.size(); i++) {
            ICrossTab crossTab = head.get(i);
            Collection<POIOperation> opCollection = crossTab.getPoiOperation();
            if (opCollection != null) {
                for (POIOperation poiOperation : opCollection) {
                    if (POIOrderOperation.OPERATION.equals(poiOperation.getOperation())) {
                        doOrderOperation(columnTreeRoot, (POIOrderOperation) poiOperation, i + 1);
                    }
                }
            }
        }
    }

    /**
     * 对树的第 <code>level</code>进行排序操作。
     * 获取第level-1层的节点，对子树排序。
     *
     * @param treeRoot       树根节点
     * @param orderOperation 排序操作
     * @param level          需要排序的层
     */
    @SuppressWarnings("unchecked")
    private void doOrderOperation(CrossTabTreeNode<CrossTabHeader> treeRoot,
                                  POIOrderOperation orderOperation, int level) {
        int i = 0;
        //排序层的上一层，也就是待排序等的父层，level-1层。得到父层后对父下面的子节点进行排序。
        TreeNode<CrossTabHeader>[] parentArray = new TreeNode[1];
        parentArray[0] = treeRoot;
        while (++i < level) {
            List<TreeNode> tmpList = new ArrayList<TreeNode>();
            for (TreeNode<CrossTabHeader> treeNode : parentArray) {
                tmpList.addAll(Arrays.asList(treeNode.getAllChild()));
            }

            parentArray = tmpList.toArray(new TreeNode[tmpList.size()]);
        }
        //对父下面的子节点进行排序，也就是第level层
        for (TreeNode<CrossTabHeader> treeNode : parentArray) {
            treeNode.orderChild(orderOperation.getOrderType());
        }
    }


    public void afterRow() throws POIException {
        // do nothing
    }

    public void beforeRow() throws POIException {
        // do nothing
    }
}
