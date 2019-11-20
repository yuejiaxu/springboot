package com.yjx.springboot.easyexcel;

import com.alibaba.excel.context.AnalysisContext;
import com.yjx.springboot.utils.easyexcel.AbstractListener;

import java.util.ArrayList;
import java.util.List;

public class UserListener extends AbstractListener<User> {

    final List<User> list = new ArrayList<User>();

    @Override
    public List<User> getDataList() {
        return list;
    }

    @Override
    public void invoke(User object, AnalysisContext context) {
        if (null != object) {
            list.add(object);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }

}
