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

    String expression() default "";

    String deExpression() default "";


}
