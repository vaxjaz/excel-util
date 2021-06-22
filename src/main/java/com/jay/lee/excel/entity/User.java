package com.jay.lee.excel.entity;

import com.jay.lee.excel.anno.ExcelName;
import com.jay.lee.excel.constant.StateEnum;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Optional;

/**
 * @Author: jay
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class User {

    @ExcelName(value = "id", order = 1)
    private Long id;

    @ExcelName(value = "姓名", required = true, order = 2)
    private String name;

    @ExcelName(value = "年龄", order = 3)
    private Integer age;

    @ExcelName(value = "生日", order = 4, expression = "method{this.parseDate(birthDay)}", deExpression = "method{this.deParseDate(birthDay)}")
    private LocalDateTime birthDay;

    @ExcelName(value = "性别", order = 5)
    private Boolean sex;

    /**
     * 0 测试状态1
     * 1 测试状态2
     * 2 测试状态3
     * 3 测试状态4
     */
    @ExcelName(value = "状态", order = 7, expression = "method{this.parseState(state)}", deExpression = "method{this.deParseState(state)}")
    private Integer state;

    @ExcelName(value = "测试", order = 6, expression = "test==1?\"就是1\":\"其他\"", deExpression = "\"其他\".equals(test)?1:0", required = true)
    private Integer test;

    public String parseState(Integer state) {
        return Optional.ofNullable(state)
                .map(StateEnum::getNameByCode)
                .orElse(null);
    }

    public Integer deParseState(String state) {

        return Optional.ofNullable(state)
                .map(StateEnum::getCodeByName)
                .orElse(null);
    }

    public String parseDate(LocalDateTime birthDay) {
        return Optional.ofNullable(birthDay)
                .map(localDateTime -> {
                    return localDateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
                })
                .orElse(null);
    }

    public LocalDateTime deParseDate(String birthDay) {
        return Optional.ofNullable(birthDay)
                .map(date -> {
                    return LocalDateTime.parse(date, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
                })
                .orElse(null);
    }


}
