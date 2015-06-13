package com.cfcc.deptone.excel.gen.inner;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ISheet;

/**
 * 自定义报表，此类报表有包含普通和交叉类型单元格
 * @author WangHuanyu
 */
public class SheetBuilderCust1 extends AbstractSheetBuilder {

	private List<BuildStep> step = new ArrayList<BuildStep>();

	protected SheetBuilderCust1(ISheet sheet) {
		super(sheet);
	}

	public void write() throws POIException {
		initStep();

		for (Iterator<BuildStep> iterator = step.iterator(); iterator.hasNext();) {
			BuildStep buildStep = iterator.next();
			buildStep.build(super.sheet);
		}
	}

	public void initStep() {
		step.add(new BuildRemoveOpSumMerger());
		step.add(new BuildRemoveFooterMerger());
		step.add(new BuildRemoveConstMerger());
		step.add(new BuildCollectPictureMerger());
		
		step.add(new BuildHeader());
		step.add(new BuildConst());
		step.add(new BuildPicture());
		step.add(new BuildFooter());

		step.add(new BuildFormula());
	}

}
