package com.jay.lee.excel.controller;

import com.jay.lee.excel.entity.Other;
import com.jay.lee.excel.entity.User;
import com.jay.lee.excel.util.ExcelUtils;
import com.jay.lee.excel.util.JsonUtils;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @Author: tomato
 * @Date: 2020/12/19 17:08
 */
@Api
@RestController
@RequestMapping("/excel")
public class ExcelController {

    @ApiOperation("导出")
    @GetMapping("/export")
    public void testExport(HttpServletResponse response) {
        List<User> list = new ArrayList<>();
        User jay = User.builder()
                .age(18)
                .birthDay(LocalDateTime.now())
                .id(1l)
                .name("JAY")
                .sex(true)
                .state(0)
                .build();
        list.add(jay);
        ExcelUtils.export(list,response,User.class,"测试");
//        ExcelUtils.buildMultiSheet(response, "test", Arrays.asList(new User()), Arrays.asList(new Other()));
    }

    @PostMapping("/upload")
    @ApiOperation("上传")
    public String upload(@RequestParam("file") MultipartFile file) {
        InputStream inputStream = null;
        try {
            inputStream = file.getInputStream();
        } catch (IOException e) {
            e.printStackTrace();
            return "";
        }
        List<User> users = ExcelUtils.readExcel(inputStream, User.class);
        System.out.println(JsonUtils.writeValueAsString(users));
        return JsonUtils.writeValueAsString(users);
    }


}
