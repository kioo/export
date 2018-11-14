package com.wangchi.demo.export.controller;


import com.wangchi.demo.export.service.PoiService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;

@RestController
public class POIController {

    @Autowired
    PoiService poiService;

    /**
     * 输出调用controller
     * @param response
     * @param dataNum
     * @return
     */
    @RequestMapping(value = "/export",produces = {"application/vnd.ms-excel;charset=UTF-8"})
    public String export(HttpServletResponse response, @RequestParam(value = "dataNum",required = false) Integer dataNum){
        return poiService.export(response,dataNum);
    }

}
