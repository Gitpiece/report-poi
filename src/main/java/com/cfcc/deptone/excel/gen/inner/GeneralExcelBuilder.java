package com.cfcc.deptone.excel.gen.inner;


import com.cfcc.deptone.excel.gen.POIException;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * 2012-3-31
 * @author thseus
 */
public class GeneralExcelBuilder extends AbstractExcelBuilder {
    //private static final Log LOGGER = LogFactory.getLog(GeneralExcelBuilder.class);

    public int build(String reportPath, Map<String, Object>[] metaDataMap, List<?>[] rptDataList) throws POIException {
        return super.build(new File(reportPath),metaDataMap,rptDataList);
    }

    public int build(String reportPath, Map<String, Object> metaDataMap, List<?> rptDataList) throws POIException {
        return super.build(new File(reportPath),metaDataMap,rptDataList);
    }
}
