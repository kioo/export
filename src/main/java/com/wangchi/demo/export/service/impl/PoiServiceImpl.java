package com.wangchi.demo.export.service.impl;


import com.wangchi.demo.export.mock.Pcd;
import com.wangchi.demo.export.service.PoiService;
import com.wangchi.demo.export.utils.DataGeneratorUtil;
import com.wangchi.demo.export.utils.ExcelUtil;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class PoiServiceImpl implements PoiService {

    /**
     * 输出方法
     *
     * @param response
     * @return
     */
    @Override
    public String export(HttpServletResponse response,int dataNum) {
        try {
            List<Pcd> listPcd = DataGeneratorUtil.generatorPcd(dataNum);
            String fileName="项目审核表";
            List<Map<String,Object>> list=createExcelRecord(listPcd);
            String columnNames[] = {"编号","名称","父节点名称","简称","级别","城市编码","区域编码"};//列名
            String keys[] = {"id","name","parentId","shortName","levelType","cityCode","zipCode"};//map中的key
            ExcelUtil.downloadWorkBook(list,keys,columnNames,fileName,response);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "excel";
    }

    /**
     * 创建Excel表中的记录
     * @param pcdList
     * @return
     */
    private List<Map<String, Object>> createExcelRecord(List<Pcd> pcdList){
        List<Map<String, Object>> listmap = new ArrayList<Map<String, Object>>();
        try {
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("sheetName", "sheet1");
            listmap.add(map);
            for (int j = 0; j < pcdList.size(); j++) {
                Pcd pcd=pcdList.get(j);
                Map<String, Object> mapValue = new HashMap<String, Object>();
                mapValue.put("id",pcd.getId());
                mapValue.put("name",pcd.getName());
                mapValue.put("parentId",pcd.getParentId());
                mapValue.put("shortName",pcd.getShortName());
                mapValue.put("levelType",pcd.getLevelType());
                mapValue.put("cityCode",pcd.getCityCode());
                mapValue.put("zipCode",pcd.getZipCode());
               //mapValue.put("submitTime", DateTimeUtil.dateToStr(projectAuditListVo.getSubmitTime(),"yyyy-MM-dd"));
               // String attachmentURL = projectAuditListVo.getAttachment()==null?"无": FileUtil.getUploadPath()+projectAuditListVo.getAttachment();

                listmap.add(mapValue);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return listmap;
    }
}
