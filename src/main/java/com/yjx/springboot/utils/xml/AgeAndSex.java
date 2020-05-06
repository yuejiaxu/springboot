package com.yjx.springboot.utils.xml;


import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

/**
 * @author yjx
 * @version V1.0
 * @description AgeAndSex
 * @since 2020/5/6 18:23
 */
@XmlRootElement(name = "AGEANDSEX")
@XmlAccessorType(XmlAccessType.FIELD)
public class AgeAndSex {
    @XmlElement(name = "AGE")
    private String age;
    @XmlElement(name = "SEX")
    private String sex;

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
