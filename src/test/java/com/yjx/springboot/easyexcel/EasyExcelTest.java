package com.yjx.springboot.easyexcel;

import com.alibaba.fastjson.JSONObject;
import com.yjx.springboot.utils.easyexcel.EasyExcelUtil;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * easyExcel测试类
 */
public class EasyExcelTest {
    @Test
    public void read() {
        long startTime = System.currentTimeMillis();   //获取开始时间
        String filePath = "D:\\bbb.xlsx";
        UserListener listener = new UserListener();
        List<User> list = EasyExcelUtil.readAllSheetRow(filePath, listener, User.class);
        System.out.println(JSONObject.toJSONString(list));
        long endTime = System.currentTimeMillis(); //获取结束时间
        System.out.println("");
        System.out.println("程序运行时间： " + (endTime - startTime) + "ms");
    }

    @Test
    public void write() throws Exception {
        long startTime = System.currentTimeMillis();   //获取开始时间
        String filePath = "D:\\bbb.xlsx";
        String sheetName = "第一个sheet";
        List<User> data = new ArrayList<User>();
        for (int i = 1; i <= 10000; i++) {
            data.add(new User(new BigDecimal(i), "name" + i, new BigDecimal(i), "adress" + i));
        }
        EasyExcelUtil.writeOneSheet(filePath, sheetName, data, User.class);
        long endTime = System.currentTimeMillis(); //获取结束时间
        System.out.println("程序运行时间： " + (endTime - startTime) + "ms");
    }

    public static void main(String[] args) throws Exception {
        String filePath = "C:\\Users\\zewe\\Desktop\\template\\test.xlsx";
        Integer sheetIndex = 1;
        Integer startRowIndex = 1;
        Integer columnIndex = 8;


    	/*String filePath1 = "C:\\Users\\zewe\\Desktop\\template\\test2.xlsx";
        String sheetName = "第一个sheet";
        List<String> headers = Arrays.asList("A列","B列","C列");
        List<List<Object>> data = new ArrayList<List<Object>>();
        List<Object> d1 = Arrays.asList(1,2,3);
        List<Object> d2 = Arrays.asList(new Date(),new Date(),new Date());
        List<Object> d3 = Arrays.asList("汉","字",null);
        data.add(d1);data.add(d2);data.add(d3);*/


        String filePath2 = "D:\\bbb.xlsx";
        String sheetName = "第一个sheet";
//        List<User> data = new ArrayList<User>();
//        for (int i = 1; i <= 110; i++) {
//            data.add(new User(new BigDecimal(i), "name" + i, new BigDecimal(i), "adress" + i));
//        }

        //List<Object> list = EasyExcelUtil.readRow(filePath, 1, 0, false);
        //OneColumnListener listener = new OneColumnListener();
        //List<OneColumn> list2 = EasyExcelUtil.readRow(filePath, sheetIndex, startRowIndex, listener, OneColumn.class);
        //List<String> list3 = EasyExcelUtil.readOneSheetOneColumn(filePath, sheetIndex,startRowIndex,columnIndex);
        //System.out.println(JSONObject.toJSONString(list3));

//        EasyExcelUtil.writeOneSheet(filePath2, sheetName, data, User.class);
//        EasyExcelUtil.writeSheets(filePath2, 20, data, User.class);
        UserListener listener = new UserListener();
//        List<User> list = EasyExcelUtil.readAllSheetRow(filePath2, listener, User.class);
        List<User> list = EasyExcelUtil.readRow(filePath2, sheetIndex, startRowIndex, listener, User.class);
        System.out.println(JSONObject.toJSONString(list));
    }
}
