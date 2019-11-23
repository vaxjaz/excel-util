package com.jay.lee.excel.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONPath;
import com.alibaba.fastjson.serializer.JSONLibDataFormatSerializer;
import com.alibaba.fastjson.serializer.SerializeConfig;
import com.alibaba.fastjson.serializer.SerializerFeature;
import lombok.extern.slf4j.Slf4j;

import java.util.Collections;
import java.util.Date;
import java.util.List;

/**
 * @Author: jay
 */
@Slf4j
public final class JsonUtils {

    private static final SerializeConfig config = new SerializeConfig();
    private static final SerializerFeature[] features;

    private JsonUtils() {
    }

    public static String writeValueAsString(Object value) {
        return value == null ? null : JSON.toJSONString(value, config, features);
    }



    public static <T> List<T> readValueAsArray(String json, Class<T> valueType) {
        return JSON.parseArray(json, valueType);
    }

    public static <T> List<T> readValueAsArray(String json, String jsonPath, Class<T> valueType) {
        Object content = JSONPath.read(json, jsonPath);
        if (content == null) {
            log.debug("path not found:{}", jsonPath);
            return Collections.emptyList();
        } else {
            return JSONArray.parseArray(content.toString(), valueType);
        }
    }

    static {
        config.put(Date.class, new JSONLibDataFormatSerializer());
        config.put(java.sql.Date.class, new JSONLibDataFormatSerializer());
        features = new SerializerFeature[]{SerializerFeature.WriteMapNullValue, SerializerFeature.WriteNullListAsEmpty, SerializerFeature.WriteNullBooleanAsFalse, SerializerFeature.WriteNullStringAsEmpty};
    }



}
