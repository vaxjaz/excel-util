package com.jay.lee.excel.entity;

import com.jay.lee.excel.anno.ExcelName;
import lombok.Data;

/**
 * @Author: tomato
 * @Date: 2020/9/18 15:50
 */
@Data
public class Other {

    @ExcelName("otherId")
    private String id;

    @ExcelName("otherName")
    private String name;

}
