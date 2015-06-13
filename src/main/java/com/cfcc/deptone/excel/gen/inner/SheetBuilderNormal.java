package com.cfcc.deptone.excel.gen.inner;

import com.cfcc.deptone.excel.gen.POIException;
import com.cfcc.deptone.excel.model.ISheet;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * 普通报表构建类
 * @author WangHuanyu
 */
public class SheetBuilderNormal extends AbstractSheetBuilder {

	private List<BuildStep> step = new ArrayList<BuildStep>();

	protected SheetBuilderNormal(ISheet sheet) {
		super(sheet);
	}

	public void write() throws POIException {
		//
		initStep();

		//
		for (Iterator<BuildStep> iterator = step.iterator(); iterator.hasNext();) {
			BuildStep type = iterator.next();
			type.build(sheet);
		}
	}

	@Override
	public void initStep() {
		//移除合并单元格
		step.add(new BuildRemoveOpSumMerger());
		step.add(new BuildRemoveFooterMerger());
		step.add(new BuildRemoveConstMerger());


		step.add(new BuildNormal());
		step.add(new BuildConst());
		step.add(new BuildOpSum());

		step.add(new BuildHeader());
		step.add(new BuildFooter());
		
		step.add(new BuildFormula());
	}

}
