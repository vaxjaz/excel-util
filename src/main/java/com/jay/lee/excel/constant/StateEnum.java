package com.jay.lee.excel.constant;

import java.util.Arrays;
import java.util.Objects;

/**
 * @Author: jay
 */
public enum StateEnum {

    TEST_STATE_1(0,"测试状态1"),
    TEST_STATE_2(1,"测试状态2"),
    TEST_STATE_3(2,"测试状态3"),
    TEST_STATE_4(3,"测试状态4")





    ;


    private Integer code;

    private String msg;

    StateEnum(Integer code, String msg) {
        this.code = code;
        this.msg = msg;
    }

    public static String getNameByCode(Integer state){
        return Arrays.stream(StateEnum.values())
                .filter(stateEnum -> Objects.equals(stateEnum.getCode(),state))
                .findFirst()
                .map(StateEnum::getMsg)
                .orElse(null);
    }

    public Integer getCode() {
        return code;
    }

    public String getMsg() {
        return msg;
    }

    public static Integer getCodeByName(String msg) {
        return Arrays.stream(StateEnum.values())
                .filter(stateEnum -> Objects.equals(stateEnum.getMsg(),msg))
                .findFirst()
                .map(StateEnum::getCode)
                .orElse(null);
    }
}
