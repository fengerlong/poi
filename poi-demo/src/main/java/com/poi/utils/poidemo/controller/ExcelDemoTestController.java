package com.poi.utils.poidemo.controller;

import com.poi.utils.poidemo.entity.WebDto;
import com.poi.utils.poidemo.utils.ExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Controller
@RequestMapping("/excel")
public class ExcelDemoTestController {

    /**
     * 测试excel以流方式导出
     * @param response
     */
    @GetMapping("/test1")
    private void test1(HttpServletResponse response){
        //构建假性数据
        List<WebDto> list = new ArrayList<WebDto>();
        list.add(new WebDto("知识林", "http://www.zslin.com", "admin", "111111", 555,new Date()));
        list.add(new WebDto("权限系统", "http://basic.zslin.com", "admin", "111111", 111,new Date()));
        list.add(new WebDto("校园网", "http://school.zslin.com", "admin", "222222", 333,new Date()));

        ExcelUtil.getInstance().exportExcelByStream(response,list,WebDto.class,"网站访问量统计表",true);
    }

    @GetMapping("/test2")
    private void test2(HttpServletRequest request) throws Exception {
        //构建假性数据
//        List<WebDto> list = new ArrayList<WebDto>();
//        list.add(new WebDto("知识林", "http://www.zslin.com", "admin", "111111", 555,new Date()));
//        list.add(new WebDto("权限系统", "http://basic.zslin.com", "admin", "111111", 111,new Date()));
//        list.add(new WebDto("校园网", "http://school.zslin.com", "admin", "222222", 333,new Date()));

        ExcelUtil.getInstance().readExcel("https://raw.githubusercontent.com/fengerlong/picture/master/%E7%BD%91%E7%AB%99%E8%AE%BF%E9%97%AE%E9%87%8F%E7%BB%9F%E8%AE%A1%E8%A1%A8.xlsx",
                request.getInputStream());
    }


}
