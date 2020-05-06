package com.yjx.springboot.utils.xml;

import javax.xml.bind.annotation.*;

/**
 * @author yjx
 * @version V1.0
 * @description Girl
 * @since 2020/5/6 18:25
 */
@XmlRootElement(name = "GIRL")
@XmlAccessorType(XmlAccessType.FIELD)
public class Girl<T> {
    @XmlElement(name = "NAME")
    private String name;
    /**
     * XmlAnyElement 这个注解可以去调生成的xml中带的xsi:type等信息，使用这个注解就不能使用 @XmlElement，
     * 需要在泛型对应的实体上增加 @XmlRootElement 注解
     * XmlElementWrapper 这个注解可以在集合外层包装一个节点
     */
    @XmlAnyElement(lax = true)
    private T ageAndSex;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public T getAgeAndSex() {
        return ageAndSex;
    }

    public void setAgeAndSex(T ageAndSex) {
        this.ageAndSex = ageAndSex;
    }
}