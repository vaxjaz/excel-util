package com.jay.lee.excel.exception;

import lombok.Data;

/**
 * @Author: jay
 */
@Data
public class ParameterException extends RuntimeException {
    public ParameterException() {
    }

    public ParameterException(String message) {
        super(message);
    }
}
