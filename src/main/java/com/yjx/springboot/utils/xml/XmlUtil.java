package com.yjx.springboot.utils.xml;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Marshaller;
import javax.xml.bind.Unmarshaller;
import java.io.StringReader;
import java.io.StringWriter;

/**
 * @ClassNAME XmlUtil
 * @Description XML与bean互相转换
 * @Author yuejiaxu
 * @Date 2020-5-6
 */
public class XmlUtil {

    /**
     * @Description bean转成xml
     * @Param [t]
     * @Return java.lang.String
     */
    public static <T> String beanToXml(T t) throws JAXBException {
        //获得 JAXBContext 类的新实例。参数为类的地址
        JAXBContext context = JAXBContext.newInstance(t.getClass());
        //创建一个可以用来将 java 内容树转换为 XML 数据的 Marshaller 对象。
        Marshaller m = context.createMarshaller();
        //创建一个StringWriter流将接收到的对象流写入xml字符串
        StringWriter sw = new StringWriter();
        //调用marshal方法进行转换
        m.marshal(t, sw);
        //将读取到的StringWriter流转成String返回
        return sw.toString();
    }

    /**
     * @Description bean转成xml(泛型使用)
     * @Param [t]
     * @Return java.lang.String
     */
    public static <T> String beanToXml(T t, Class c) throws JAXBException {
        //获得 JAXBContext 类的新实例。参数为类的地址
        JAXBContext context = JAXBContext.newInstance(t.getClass(), c);
        //创建一个可以用来将 java 内容树转换为 XML 数据的 Marshaller 对象。
        Marshaller m = context.createMarshaller();
        //创建一个StringWriter流将接收到的对象流写入xml字符串
        StringWriter sw = new StringWriter();
        //调用marshal方法进行转换
        m.marshal(t, sw);
        //将读取到的StringWriter流转成String返回
        return sw.toString();
    }

    /**
     * @Description xml 转成 bean
     * @Param [xml, t]
     * @Return T
     */
    public static <T> T xmlToBean(String xml, T t) throws JAXBException {
        ////获得 JAXBContext 类的新实例。参数为类的地址
        JAXBContext context = JAXBContext.newInstance(t.getClass());
        //创建一个可以用来将 XML 数据转换为 java 内容树的 Unmarshaller 对象。
        Unmarshaller um = context.createUnmarshaller();
        //创建一个StringReader将xml报文转成流
        StringReader sr = new StringReader(xml);
        //调用unmarshal进行转换，并把Object类型强转为调用者的类型
        t = (T) um.unmarshal(sr);
        //将对象返回给调用者
        return t;
    }

    /**
     * @Description xml 转成 bean(泛型使用)
     * @Param [xml, t]
     * @Return T
     */
    public static <T> T xmlToBean(String xml, T t, Class c) throws JAXBException {
        ////获得 JAXBContext 类的新实例。参数为类的地址
        JAXBContext context = JAXBContext.newInstance(t.getClass(), c);
        //创建一个可以用来将 XML 数据转换为 java 内容树的 Unmarshaller 对象。
        Unmarshaller um = context.createUnmarshaller();
        //创建一个StringReader将xml报文转成流
        StringReader sr = new StringReader(xml);
        //调用unmarshal进行转换，并把Object类型强转为调用者的类型
        t = (T) um.unmarshal(sr);
        //将对象返回给调用者
        return t;
    }

    public static void main(String[] args) {
        Girl girl = new Girl();
        girl.setName("小红");
        AgeAndSex ageAndSex = new AgeAndSex();
        ageAndSex.setAge("18");
        ageAndSex.setSex("女");
        girl.setAgeAndSex(ageAndSex);
        System.out.println("组装完成的对象值为：" + girl);
        //方法调用并输出
        Girl newGirl = new Girl();
        String xml;
        try {
            xml = XmlUtil.beanToXml(girl, AgeAndSex.class);
            System.out.println("bean转成xml格式为：" + xml);
            newGirl = XmlUtil.xmlToBean(xml, newGirl, AgeAndSex.class);
            System.out.println("xml转成bean格式为：" + newGirl);
        } catch (JAXBException e) {
            e.printStackTrace();
        }

    }
}
