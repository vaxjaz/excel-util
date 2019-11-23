package com.jay.lee.excel.exception;

import lombok.Data;

/**
 * @Author: jay
 */
@Data
public class NotFoundException extends RuntimeException {

    public NotFoundException() {
    }

    public NotFoundException(String message) {
        super(message);
    }
}
