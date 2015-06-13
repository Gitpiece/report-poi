package com.cfcc.deptone.excel;

import com.cfcc.deptone.excel.data.TempObj;
import com.cfcc.deptone.excel.data.TestData;
import com.cfcc.deptone.excel.gen.ExcelBuilder;
import com.cfcc.deptone.excel.gen.ExcelBuilderFactory;
import com.cfcc.deptone.excel.model.ICrossTab;
import com.cfcc.deptone.excel.poi.support.CrossTabData;
import com.cfcc.deptone.excel.poi.support.CrossTabHeader;
import com.cfcc.deptone.excel.util.POIExcelUtil;
import com.cfcc.deptone.excel.util.TreeNode;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;

import java.io.File;
import java.lang.reflect.Method;
import java.util.*;

public class CrossTest {
    Map<String, Method> methodMap = new HashMap<String, Method>();

    @Test
    public void testCross2003() throws Exception {
        FileUtils.copyFile(new File(TestData.resourcePath + "cross.xls"), new File(TestData.reportPath + "cross-out.xls"));
        ExcelBuilderFactory.getBuilder().build(TestData.reportPath + "cross-out.xls", TestData.getMetadata(), TestData.getCrossList());
    }

    @Test
    public void testCross2007() throws Exception {
        FileUtils.copyFile(new File(TestData.resourcePath + "cross.xlsx"), new File(TestData.reportPath + "cross-out.xlsx"));
        ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
        excelBuilder.build(TestData.reportPath + "cross-out.xlsx", TestData.getMetadata(), TestData.getCrossList());
    }

    @Test
    public void testCross2007_2() throws Exception {
        FileUtils.copyFile(new File(TestData.resourcePath + "cross_2.xlsx"), new File(TestData.reportPath + "cross_2-out.xlsx"));
        ExcelBuilder excelBuilder = ExcelBuilderFactory.getBuilder();
        excelBuilder.build(TestData.reportPath + "cross_2-out.xlsx", TestData.getMetadata(), TestData.getCrossList());
    }

    List<POICrossTab> colomnhead;
    List<POICrossTab> rowhead;
    TreeNode<CrossTabHeader> columntree;
    TreeNode<CrossTabHeader> rowtree;

    /**
     * 测试生成交叉报表的列表头
     */
    @Test
    public void testTreeColumn() {
        POICrossTab head1 = new POICrossTab("title1");
        POICrossTab head2 = new POICrossTab("title2");
        POICrossTab head3 = new POICrossTab("title3");
        POICrossTab name = new POICrossTab("name");
        POICrossTab sbtcode = new POICrossTab("sbtcode");

        colomnhead = new ArrayList<POICrossTab>(3);
        colomnhead.add(head1);
        colomnhead.add(head2);
        colomnhead.add(head3);
        rowhead = new ArrayList<POICrossTab>(3);
        rowhead.add(name);
        rowhead.add(sbtcode);

        List<TempObj> listdata = TestData.getCrossList();

        //整理表头信息
        columntree = buildCrossTabColumn(colomnhead, listdata);
        rowtree = buildCrossTabColumn(rowhead, listdata);
        System.out.println(leafCount(columntree));
        System.out.println(leafCount(columntree.getChild(0)));
        System.out.println(leafCount(columntree.getChild(0).getChild(0)));
        System.out.println(leafCount(columntree.getChild(0).getChild(1)));
        System.out.println(leafCount(columntree.getChild(0).getChild(0).getChild(0)));
        int row = leafCount(rowtree), column = leafCount(columntree);
        System.out.println(row);
        System.out.println(column);
        CrossTabData[][] dataMatrix = new CrossTabData[row][column];

        fillDataMatrix(dataMatrix, listdata);


        for (int i = 0; i < dataMatrix.length; i++) {
            CrossTabData[] crossTabDatas = dataMatrix[i];
            for (int j = 0; j < crossTabDatas.length; j++) {
                CrossTabData crossTabData = crossTabDatas[j];
                if (crossTabData != null)
                    System.out.print("\t" + ((TempObj) crossTabData.getOriginalObj()).getEndamt());
                else
                    System.out.print("\t  null");
            }
            System.out.println();
        }
    }

    /**
     * 把数据填入数据矩阵。
     * 通过构建好的表头定位数据坐标，构建数据矩阵。
     *
     * @param dataMatrix 数据矩阵
     * @param data       数据集合
     */
    private void fillDataMatrix(CrossTabData[][] dataMatrix, List<?> data) {
        for (Iterator iterator = data.iterator(); iterator.hasNext(); ) {
            Object object = (Object) iterator.next();
            Point point = computePoint(object);
            CrossTabData crossTabData = new CrossTabData();
            crossTabData.setOriginalObj(object);
            dataMatrix[point.x][point.y] = crossTabData;
        }
    }

    /**
     * 计算数据坐标
     *
     * @param object 数据对象
     * @return
     */
    private Point computePoint(Object object) {
        StringBuilder columnStr = new StringBuilder();
        for (Iterator iterator = colomnhead.iterator(); iterator.hasNext(); ) {
            POICrossTab crossTab = (POICrossTab) iterator.next();
            columnStr.append(String.valueOf(POIExcelUtil.getPropertyValue(object, crossTab.getPropertyName(), methodMap)));
            columnStr.append(">");
        }
        StringBuilder rowStr = new StringBuilder();
        for (Iterator iterator = rowhead.iterator(); iterator.hasNext(); ) {
            POICrossTab crossTab = (POICrossTab) iterator.next();
            rowStr.append(String.valueOf(POIExcelUtil.getPropertyValue(object, crossTab.getPropertyName(), methodMap)));
            rowStr.append(">");
        }
        //rowtree决定x坐标，columntree决定y坐标。
        return new Point(computeTreeCoordinate(rowtree, rowStr), computeTreeCoordinate(columntree, columnStr));
    }

    /**
     * 计算树坐标
     *
     * @return
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
                if (child.length == 0) break;
                i = 0;//坐标清零
                pathArrIndex++;
            }
            //如果节点是叶子节点，左叶子加1
            else if (!child[0].hasChild()) {
                leftLeafCount += i;
            }
            //否则，获取当前节点的叶子节点个数
            else {
                leftLeafCount += this.leafCount(treeNode);
            }
        }

        return leftLeafCount;
    }

    /**
     * 坐标类
     *
     * @author wanghuanyu
     */
    private class Point {
        int x, y;

        protected Point(int x, int y) {
            this.x = x;
            this.y = y;
        }
    }

    /**
     * @param root
     * @return
     */
    private int leafCount(TreeNode<CrossTabHeader> root) {
        TreeNode<CrossTabHeader>[] childs = root.getAllChild();
        int count = 0;
        for (int i = 0; i < childs.length; i++) {
            TreeNode<CrossTabHeader> treeNode = childs[i];
            if (treeNode != null && treeNode.getAllChild().length > 0) {
                count += leafCount(treeNode);
            } else {
                count += childs.length;
                break;
            }
        }

        return count;
    }

    private void printTree(TreeNode<CrossTabHeader> tree) {
        TreeNode<CrossTabHeader>[] trees = tree.getAllChild();
        for (int i = 0; i < trees.length; i++) {
            TreeNode<CrossTabHeader> treeNode = trees[i];
            System.out.print(treeNode.getLabel() + ",");
        }
        System.out.println();
        for (int i = 0; i < trees.length; i++) {
            printTree(trees[i]);
        }
    }

    /**
     * 整理表头信息
     *
     * @param head
     * @param listdata
     */
    private TreeNode<CrossTabHeader> buildCrossTabColumn(List<POICrossTab> head, List<TempObj> listdata) {
        TreeNode<CrossTabHeader> ColumnTreeRoot = new TreeNode<CrossTabHeader>();

        //循环数据
        for (Iterator<TempObj> iterator = listdata.iterator(); iterator.hasNext(); ) {
            TempObj dataObj = (TempObj) iterator.next();

            //每一个数据对象，通过配置的头信息创建一个树的分支
            TreeNode<CrossTabHeader> currentTreeParent = ColumnTreeRoot;
            for (Iterator<POICrossTab> iterator2 = head.iterator(); iterator2.hasNext(); ) {
                POICrossTab crossTab = iterator2.next();

                Object value = POIExcelUtil.getPropertyValue(dataObj, crossTab.getPropertyName(), methodMap);
                //如果当前节点不包含 label 是 value的节点，创建子节点，并设置currentTreeParent为currentTreeNode
                if (!currentTreeParent.containNode(String.valueOf(value))) {
                    currentTreeParent = new TreeNode<CrossTabHeader>(currentTreeParent, String.valueOf(value),
                            new CrossTabHeader(crossTab.getPropertyName(),
                                    String.valueOf(value)));
                }
                //否则，设置 currentTreeParent 为 label 是 value的子节点
                else {
                    currentTreeParent = currentTreeParent.getChild(String.valueOf(value));
                }

            }
        }

        return ColumnTreeRoot;
    }

    private class POICrossTab {
        protected String crossType;
        private int level = 0;
        private String propertyName;

        protected POICrossTab(String pn) {
            propertyName = pn;
        }

        public String getPropertyName() {
            return propertyName;
        }

        public boolean isColumn() {
            return ICrossTab.CROSSTAB_COLUMN.equals(crossType);
        }

        public boolean isRow() {
            return ICrossTab.CROSSTAB_ROW.equals(crossType);
        }

        public boolean isData() {
            return ICrossTab.CROSSTAB_DATA.equals(crossType);
        }

        public int level() {
            return level;
        }
    }

}
