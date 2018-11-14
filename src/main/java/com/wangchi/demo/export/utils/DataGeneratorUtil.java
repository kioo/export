package com.wangchi.demo.export.utils;

import com.wangchi.demo.export.mock.Pcd;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * 数据生成工具，用来生成模拟的数据
 */
public class DataGeneratorUtil {

    /**
     * 数据生成方法
     * @param dataNumber
     * @return
     */
    public static List<Pcd> generatorPcd(Integer dataNumber) {
        if (dataNumber==0){ // 如果没有传入dataNumber 设置默认值为100
            dataNumber=100;
        }
        List<Pcd> listPcd = new ArrayList<>();
        int id = 10000;
        int cityCode = 010;
        Random randomNum = new Random(); // 随机数生成对象，用来生成随机数
        for (int i = 0; i < dataNumber; i++) {
            Pcd pcd = new Pcd();
            pcd.setId(id+i);
            pcd.setName("北京市"+i);
            pcd.setParentId(randomNum.nextInt());
            pcd.setShortName("北京"+i);
            pcd.setLevelType(i+"");
            pcd.setCityCode(cityCode+i+"");
            pcd.setZipCode(id+i+"");
            listPcd.add(pcd);
        }
        return listPcd;
    }
}
