package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ISheet;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * 交叉报表构建类
 * @author WangHuanyu
 *
 */
public class SheetBuilderCrossTab extends AbstractSheetBuilder {

	private List<BuildStep> step = new ArrayList<BuildStep>();

	protected SheetBuilderCrossTab(ISheet sheet) {
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
		//移除合并单元格
		step.add(new BuildRemoveOpSumMerger());
		step.add(new BuildRemoveHeaderMerger());
		step.add(new BuildRemoveFooterMerger());
		step.add(new BuildRemoveConstMerger());

		step.add(new BuildCrossTab());
		step.add(new BuildCrossTabSum());
		step.add(new BuildOpSum());
		step.add(new BuildConst());
		step.add(new BuildHeader());
		step.add(new BuildFooter());

		step.add(new BuildFormula());
	}

}
