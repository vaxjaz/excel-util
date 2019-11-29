package com.jay.lee.excel.util;

import com.jay.lee.excel.anno.ExcelName;
import com.jay.lee.excel.exception.NotFoundException;
import com.jay.lee.excel.exception.ParameterException;
import org.apache.commons.jexl2.Expression;
import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.JexlEngine;
import org.apache.commons.jexl2.MapContext;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @Author: jay
 */
public final class ExcelUtils {

    private ExcelUtils() {
    }

    public static void export(List<? extends Object> list, HttpServletResponse response, Class<? extends Object> clzz, String name) {
        SXSSFWorkbook workbook;
        if (!CollectionUtils.isEmpty(list)) {
            workbook = new SXSSFWorkbook(list.size());
        } else {
            workbook = new SXSSFWorkbook();
        }
        SXSSFSheet sheet = workbook.createSheet(name + ".xls");
        AtomicInteger rowIndex = new AtomicInteger(0);
        createHeader(workbook, sheet, clzz);
        for (Object o : list) {
            // 分页
//            if (rowIndex.get() == 50) {
//                sheet = workbook.createSheet();
//                createHeader(workbook, sheet, clzz);
//                rowIndex = new AtomicInteger(0);
//            }
            SXSSFRow row = sheet.createRow(rowIndex.incrementAndGet());
            int columnIndex = 0;
            for (Class<?> clss = clzz; clss != Object.class; clss = clss.getSuperclass()) {
                Field[] fields = clss.getDeclaredFields();
                for (int i = 0; i < fields.length; i++) {
                    Field field = fields[i];
                    ExcelName annotation = field.getAnnotation(ExcelName.class);
                    if (null != annotation) {
                        SXSSFCell cell = row.createCell(columnIndex);
                        ++columnIndex;
                        try {
                            String fieldName = field.getName();
                            PropertyDescriptor propertyDescriptor = new PropertyDescriptor(fieldName, o.getClass());
                            Method readMethod = propertyDescriptor.getReadMethod();
                            Object invoke = readMethod.invoke(o);
                            String expression = annotation.expression();
                            if (StringUtils.hasText(expression)) {
                                if (expression.startsWith("method")) {
                                    invoke = eval(expression, fieldName, invoke, clzz);
                                } else {
                                    invoke = eval(expression, fieldName, invoke);
                                }
                            }
                            if (null != invoke) {
                                if (invoke instanceof LocalDateTime){
                                    invoke = DateTimeUtils.date2Str((LocalDateTime) invoke);
                                }
                                cell.setCellValue(invoke.toString());
                            } else {
                                cell.setCellValue("");
                            }
                        } catch (IntrospectionException | IllegalAccessException | InvocationTargetException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        }
        pwrite(response, workbook, name);
    }

    private static Object eval(String expression, String name, Object value, Class<?> clzz) {
        expression = expression.replace("method", "").replace("{", "").replace("}", "");
        JexlEngine jexl = new JexlEngine();
        Expression e = jexl.createExpression(expression);
        JexlContext jc = new MapContext();
        Object o = null;
        try {
            o = clzz.newInstance();
        } catch (InstantiationException | IllegalAccessException e1) {
            e1.printStackTrace();
        }
        jc.set("this", o);
        jc.set(name, value);
        return e.evaluate(jc);
    }

    private static Object eval(String expression, String name, Object value) {
        JexlEngine jexl = new JexlEngine();
        Expression e = jexl.createExpression(expression);
        JexlContext jc = new MapContext();
        jc.set(name, value);
        return e.evaluate(jc);
    }

    private static void createHeader(SXSSFWorkbook workbook, SXSSFSheet sheet, Class first) {
        sheet.trackAllColumnsForAutoSizing();
        SXSSFRow head = sheet.createRow(0);
        // 设置单元格为文本
        CellStyle cellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("@"));
        int headerSize = 0;
        for (Class<?> clzz = first; clzz != Object.class; clzz = clzz.getSuperclass()) {
            Field[] declaredFields = clzz.getDeclaredFields();
            for (int i = 0; i < declaredFields.length; i++) {
                Field declaredField = declaredFields[i];
                ExcelName excelName = declaredField.getAnnotation(ExcelName.class);
                if (null != excelName) {
                    String value = excelName.value();
                    SXSSFCell cell = head.createCell(headerSize);
                    // 设置列宽
                    sheet.autoSizeColumn(headerSize);
                    cell.setCellValue(value);
                    setCellStyle(cell, workbook, excelName.required(), cellStyle);
                    sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 17 / 10);
                    ++headerSize;
                }

            }
        }

    }

    public static String encodeDownloadFileName(HttpServletRequest request, String filename) {
        try {
            String agent = request.getHeader("USER-AGENT");
            if (StringUtils.isEmpty(agent)) {
                return filename;
            }
            if (agent.contains("Firefox")) {//Firefox
                filename = "=?UTF-8?B?" + (new String(Base64.getEncoder().encode(filename.getBytes(StandardCharsets.UTF_8)))) + "?=";
            } else if (agent.contains("Chrome")) {//Chrome
                filename = new String(filename.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1);
            } else {//IE7+
                filename = URLEncoder.encode(filename, "UTF-8");
                filename = filename.replace("+", "%20");
            }
        } catch (Throwable e) {
        }
        return filename;
    }

    private static void pwrite(HttpServletResponse response, Workbook workbook, String fileName) {
        response.setCharacterEncoding("UTF-8");
//        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setContentType("multipart/form-data");
        try {
            response.addHeader("Content-Disposition", "attachment; filename=" + new String(fileName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1) + ".xls");
        } catch (Exception e) {
            e.printStackTrace();
            fileName = UUID.randomUUID().toString() + ".xls";
            response.addHeader("Content-Disposition", "attachment; filename=" + fileName + ".xls");
        }
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void setCellStyle(Cell cell, SXSSFWorkbook sxssfWorkbook, boolean required, CellStyle cellStyle) {
        if (required) {
            // 设置字体
            CellStyle requireStyle = sxssfWorkbook.createCellStyle();
            Font requiredFont = sxssfWorkbook.createFont();
            requiredFont.setColor(Font.COLOR_RED);
            // 设置单元格格式
            requiredFont.setFontName("宋体");
            // 字体加粗
            requiredFont.setBold(true);
            requiredFont.setFontHeightInPoints((short) 13);
            requireStyle.setFont(requiredFont);
            requireStyle.setDataFormat(sxssfWorkbook.createDataFormat().getFormat("@"));
            cell.setCellStyle(requireStyle);
            return;
        }
        // 设置字体
        Font font = sxssfWorkbook.createFont();
        // 设置单元格格式
        font.setFontName("宋体");
        // 字体加粗
        font.setBold(true);
        font.setFontHeightInPoints((short) 13);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }

    public static <T> List<T> readExcel(InputStream in, Class<T> clzz) {
        List<T> list = new ArrayList<>();
        XSSFWorkbook sheets = null;
        try {
            sheets = new XSSFWorkbook(in);
            XSSFSheet sheetAt = sheets.getSheetAt(0);
            if (null == sheetAt) {
                throw new NotFoundException("未找到对应的excle文件");
            }
            short cellNum = sheetAt.getRow(0).getLastCellNum();
            Map<Integer, String> header = new HashMap<>(9);
            AtomicInteger integer = new AtomicInteger();
            for (int i = 0; i <= sheetAt.getLastRowNum(); i++) {
                XSSFRow row = sheetAt.getRow(i);
                integer.incrementAndGet();
                T t = BeanUtils.instantiateClass(clzz);
                BeanWrapperImpl beanWrapper = new BeanWrapperImpl(t);
                for (short j = 0; j < cellNum; j++) {
                    String cellValue = Optional.ofNullable(row.getCell(j))
                            .map(XSSFCell::toString)
                            .orElse(null);
                    if (i == 0) {
                        header.put((int) j, cellValue);
                        continue;
                    }
                    String name = header.get((int) j);
                    Arrays.stream(clzz.getDeclaredFields())
                            .filter(field -> {
                                ExcelName annotation = field.getAnnotation(ExcelName.class);
                                if (null == annotation) {
                                    return false;
                                } else {
                                    return annotation.value().equals(name);
                                }
                            })
                            .findFirst()
                            .ifPresent(field -> {
                                ExcelName annotation = field.getAnnotation(ExcelName.class);
                                if (annotation.required() && StringUtils.isEmpty(cellValue)) {
                                    throw new ParameterException("第" + integer.get() + "行" + annotation.value() + "不能为空");
                                }
                                field.setAccessible(true);
                                try {
                                    String s = annotation.deExpression();
                                    if (StringUtils.hasText(s)) {
                                        Object eval;
                                        if (s.startsWith("method")) {
                                            eval = eval(s, field.getName(), cellValue, clzz);
                                        } else {
                                            eval = eval(s, field.getName(), cellValue);
                                        }
                                        if (null != eval) {
                                            field.set(t, eval);
                                        }
                                    } else {
                                        Object o = beanWrapper.convertForProperty(cellValue, field.getName());
                                        field.set(t, o);
                                    }
                                } catch (IllegalAccessException e) {
                                    e.printStackTrace();
                                }
                            });
                }
                if (i != 0) {
                    // 排除第一行表头数据
                    list.add(t);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (null != sheets) {
                    sheets.close();
                }
                in.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return list;
    }

}
