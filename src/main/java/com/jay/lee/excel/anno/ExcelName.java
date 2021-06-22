package com.jay.lee.excel.anno;

import java.lang.annotation.*;

/**
 * @Author: jay
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelName {

    String value() default "";

    boolean required() default false;

    int validLen() default -1;

    /**
     * 动态导出解析表达式
     * example:
     * method{this.functionName(filedName)}
     * functionName 在当前类中实现对应逻辑方法
     *
     * @return
     */
    String expression() default "";

    /**
     * 导入动态反解析表达式
     * 使用方法同 expression
     *
     * @return
     */
    String deExpression() default "";

    /**
     * 数字保留几位小数
     *
     * @return
     */
    int numberScale() default 0;

    /**
     * 排序
     *
     * @return
     */
    int order() default 0;


}
