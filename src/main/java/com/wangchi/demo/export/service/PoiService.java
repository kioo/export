package com.wangchi.demo.export.service;

import javax.servlet.http.HttpServletResponse;

public interface PoiService {
    String export(HttpServletResponse response,int dataNum);
}
