package com.yjx.springboot.easyexcel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

import java.math.BigDecimal;

public class User extends BaseRowModel {

    @ExcelProperty(value = "编号", index = 0)
    private BigDecimal id;

    @ExcelProperty(value = "姓名", index = 1)
    private String name;

    @ExcelProperty(value = "年龄", index = 2)
    private BigDecimal age;  // 直接使用 Integer 解析错误  1 -> 1.0 ?

    @ExcelProperty(value = "地址", index = 3)
    private String adress;

    public BigDecimal getId() {
        return id;
    }

    public void setId(BigDecimal id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public BigDecimal getAge() {
        return age;
    }

    public void setAge(BigDecimal age) {
        this.age = age;
    }

    public String getAdress() {
        return adress;
    }

    public void setAdress(String adress) {
        this.adress = adress;
    }

    public User() {
        super();
    }

    public User(BigDecimal id, String name, BigDecimal age, String adress) {
        super();
        this.id = id;
        this.name = name;
        this.age = age;
        this.adress = adress;
    }

}
