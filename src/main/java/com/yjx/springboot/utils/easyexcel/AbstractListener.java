package com.yjx.springboot.utils.easyexcel;

import com.alibaba.excel.event.AnalysisEventListener;

import java.util.List;

/**
 * extends AnalysisEventListener add getDataList
 *
 * @param <T>
 * @author ZeWe
 */
public abstract class AbstractListener<T> extends AnalysisEventListener<T> {

    public abstract List<T> getDataList();


}
