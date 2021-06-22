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
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.beans.PropertyEditorSupport;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * @Author: jay
 */
public final class ExcelUtils {

    private ExcelUtils() {
    }

    private static final Pattern method_regex = Pattern.compile("^method\\{(this\\.\\w+?\\(\\w*\\))\\}$");


    public static void buildMultiSheet(HttpServletResponse response, String name, List<? extends Object>... list) {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        SXSSFSheet sheet = null;
        SXSSFRow row = null;
        int pageIndex = 0;
        for (List<?> item : list) {
            int pageRowNo = 0;
            sheet = workbook.createSheet(String.format("%s%s%s%s", name, "-", pageIndex, ".xls"));
            sheet = workbook.getSheetAt(pageIndex);
            for (int i = 0; i < item.size(); i++) {
                int columnIndex = 0;
                Object o = item.get(i);
                row = sheet.createRow(++pageRowNo);
                createHeader(sheet, o.getClass(), workbook);
                buildCell(o.getClass(), row, o, columnIndex);
            }
            ++pageIndex;
        }
        pwrite(response, workbook, name);
    }


    public static void exportBigData(List<? extends Object> list, HttpServletResponse response, Class<? extends Object> clzz, String name) {
        SXSSFWorkbook workbook = new SXSSFWorkbook(10000);
        SXSSFSheet sheet = null;
        SXSSFRow row;
        int rowNum = 0;
        int pageRowNo = 0;
        for (Object o : list) {
            int rowSiz = rowNum % 5000;
            int sheetIndex = rowNum / 5000;
            if (rowSiz == 0) {
                sheet = workbook.createSheet(name + sheetIndex + ".xlsx");
                sheet = workbook.getSheetAt(sheetIndex);
                createHeader(sheet, clzz, workbook);
                pageRowNo = 0;
            }
            rowNum++;
            row = sheet.createRow(++pageRowNo);
            int columnIndex = 0;
            buildCell(clzz, row, o, columnIndex);
        }
        pwrite(response, workbook, name);

    }

    private static void buildCell(Class<?> clzz, SXSSFRow row, Object o, int columnIndex) {
        Cell cell;
        List<Field> fieldList = getFields(clzz);
        for (Field field : fieldList) {
            ExcelName annotation = field.getAnnotation(ExcelName.class);
            if (Objects.nonNull(annotation)) {
                cell = row.createCell(columnIndex);
                ++columnIndex;
                try {
                    String fieldName = field.getName();
                    PropertyDescriptor propertyDescriptor = new PropertyDescriptor(fieldName, o.getClass());
                    Method readMethod = propertyDescriptor.getReadMethod();
                    Object invoke = readMethod.invoke(o);
                    String expression = annotation.expression();
                    if (StringUtils.hasText(expression)) {
                        Matcher matcher = method_regex.matcher(expression);
                        if (matcher.find()) {
                            invoke = eval(matcher.group(1), fieldName, invoke, o);
                        } else {
                            invoke = eval(expression, fieldName, invoke);
                        }
                    }
                    if (null != invoke) {
                        cell.setCellValue(data2String(invoke));
                    } else {
                        cell.setCellValue("");
                    }
                } catch (IntrospectionException | IllegalAccessException | InvocationTargetException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private static List<Field> getFields(Class<?> clzz) {
        List<Field> fieldList = new ArrayList<>();
        for (Class<?> clss = clzz; clss != Object.class; clss = clss.getSuperclass()) {
            Field[] fields = clss.getDeclaredFields();
            fieldList.addAll(Arrays.asList(fields));
        }
        fieldList = sortField(fieldList);
        return fieldList;
    }

    private static String data2String(Object invoke) {
        if (Objects.isNull(invoke)) {
            return null;
        }
        if (invoke instanceof LocalDateTime) {
            invoke = ((LocalDateTime) invoke).format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        }
        if (invoke instanceof LocalDate) {
            invoke = ((LocalDate) invoke).format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        }
        if (invoke instanceof Date) {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            invoke = simpleDateFormat.format(invoke);
        }
        return invoke.toString();
    }

    private static void registerCustomEditor(BeanWrapperImpl beanWrapper) {
        beanWrapper.registerCustomEditor(LocalDateTime.class, new PropertyEditorSupport() {
            @Override
            public void setAsText(String text) throws IllegalArgumentException {
                if (!StringUtils.isEmpty(text)) {
                    this.setValue(LocalDateTime.parse(text, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
                }
            }
        });
        beanWrapper.registerCustomEditor(LocalDate.class, new PropertyEditorSupport() {
            @Override
            public void setAsText(String text) throws IllegalArgumentException {
                if (!StringUtils.isEmpty(text)) {
                    this.setValue(LocalDate.parse(text, DateTimeFormatter.ofPattern("yyyy-MM-dd")));
                }
            }
        });
        beanWrapper.registerCustomEditor(Date.class, new PropertyEditorSupport() {
            @Override
            public void setAsText(String text) throws IllegalArgumentException {
                if (!StringUtils.isEmpty(text)) {
                    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    try {
                        this.setValue(dateFormat.parse(text));
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }
                }
            }
        });
    }


    public static void export(List<? extends Object> list, HttpServletResponse response, Class<? extends Object> clzz, String name) {
        if (list.size() >= 10000) {
            exportBigData(list, response, clzz, name);
            return;
        }
        SXSSFWorkbook workbook;
        if (!CollectionUtils.isEmpty(list)) {
            workbook = new SXSSFWorkbook(list.size());
        } else {
            workbook = new SXSSFWorkbook();
        }
        SXSSFSheet sheet = workbook.createSheet(name + ".xlsx");
        AtomicInteger rowIndex = new AtomicInteger(0);
        createHeader(sheet, clzz, workbook);
        for (Object o : list) {
            SXSSFRow row = sheet.createRow(rowIndex.incrementAndGet());
            int columnIndex = 0;
            buildCell(clzz, row, o, columnIndex);
        }
        pwrite(response, workbook, name);
    }

    private static void createHeader(SXSSFSheet sheet, Class<?> first, SXSSFWorkbook workbook) {
        SXSSFRow head = sheet.createRow(0);
        sheet.setDefaultColumnWidth((short) 30);
        // 设置单元格为文本
        int headerSize = 0;
        List<Field> fields = getFields(first);
        for (int i = 0; i < fields.size(); i++) {
            Field declaredField = fields.get(i);
            ExcelName excelName = declaredField.getAnnotation(ExcelName.class);
            if (null != excelName) {
                String value = excelName.value();
                SXSSFCell cell = head.createCell(headerSize);
                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                // 字体加粗
                font.setBold(true);
                font.setFontHeightInPoints((short) 13);
                style.setFont(font);
                style.setDataFormat(workbook.createDataFormat().getFormat("@"));
                cell.setCellStyle(style);
                sheet.setColumnWidth(i, 3000);
                // 设置单元格格式
                font.setFontName("宋体");
                if (excelName.required()) {
                    // 设置字体
                    font.setColor(Font.COLOR_RED);
                }
                cell.setCellValue(value);
                ++headerSize;
            }

        }
    }

    private static List<Field> sortField(List<Field> fields) {
        return fields.stream()
                .sorted((f1, f2) -> {
                    ExcelName excelName1 = f1.getAnnotation(ExcelName.class);
                    ExcelName excelName2 = f2.getAnnotation(ExcelName.class);
                    int o1 = 0;
                    int o2 = 0;
                    if (null != excelName1) {
                        o1 = excelName1.order();
                    }
                    if (null != excelName2) {
                        o2 = excelName2.order();
                    }
                    return o1 - o2;
                })
                .collect(Collectors.toList());
    }


    private static Object eval(String expression, String name, Object value, Object object) {
        JexlEngine jexl = new JexlEngine();
        Expression e = jexl.createExpression(expression);
        JexlContext jc = new MapContext();
        jc.set("this", object);
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
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
//        response.setContentType("multipart/form-result");
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

    private static <T> void convertSheet(XSSFSheet sheetAt, Class<T> clzz, List<T> list) {
        short cellNum = sheetAt.getRow(0).getLastCellNum();
        Map<Integer, String> header = new HashMap<>(9);
        AtomicInteger integer = new AtomicInteger();
        for (int i = 0; i <= sheetAt.getLastRowNum(); i++) {
            XSSFRow row = sheetAt.getRow(i);
            if (null == row || checkRowIsEmpty(row)) {
                break;
            }
            integer.incrementAndGet();
            T t = BeanUtils.instantiateClass(clzz);
            BeanWrapperImpl beanWrapper = new BeanWrapperImpl(t);
            for (short j = 0; j < cellNum; j++) {
                XSSFCell cell = row.getCell(j);
                CellType cellType = null;
                if (Objects.nonNull(cell)) {
                    cellType = cell.getCellType();
                }
                final String[] cellValue = {Optional.ofNullable(cell)
                        .map(ExcelUtils::getCellValue)
                        .orElse(null)};
                if (i == 0) {
                    // 头信息
                    header.put((int) j, cellValue[0]);
                    continue;
                }
                String name = header.get((int) j);
                CellType finalCellType = cellType;
                Arrays.stream(clzz.getDeclaredFields())
                        .filter(field -> existsExcelAnno(name, field))
                        .findFirst()
                        .ifPresent(field -> {
                            ExcelName annotation = field.getAnnotation(ExcelName.class);
                            validateValue(integer, cellValue[0], annotation);
                            if (Objects.nonNull(finalCellType) && finalCellType == CellType.NUMERIC) {
                                BigDecimal bigDecimal = null;
                                try {
                                    bigDecimal = new BigDecimal(cellValue[0]).setScale(annotation.numberScale(), RoundingMode.HALF_UP);
                                    cellValue[0] = bigDecimal.toString();
                                } catch (Exception e) {
                                }
                            }
                            try {
                                invokeValue(clzz, t, beanWrapper, cellValue[0], field, annotation);
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
    }


    public static <T> List<T> readExcel(InputStream in, Class<T> clzz) {
        List<T> list = new ArrayList<>();
        XSSFWorkbook sheets = null;
        try {
            sheets = new XSSFWorkbook(in);
            XSSFSheet sheetAt = sheets.getSheetAt(0);
            if (null == sheetAt) {
                throw new NotFoundException("未找到对应的sheet文件");
            }
            convertSheet(sheetAt, clzz, list);
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

    private static <T> void invokeValue(Class<T> clzz, T t, BeanWrapperImpl beanWrapper, String cellValue, Field field, ExcelName annotation) throws IllegalAccessException {
        String expression = annotation.deExpression();
        Object eval;
        if (StringUtils.hasText(expression)) {
            Matcher matcher = method_regex.matcher(expression);
            if (matcher.find()) {
                eval = eval(matcher.group(1), field.getName(), cellValue, clzz);
            } else {
                eval = eval(expression, field.getName(), cellValue);
            }
        } else {
            eval = beanWrapper.convertForProperty(cellValue, field.getName());
        }
        if (eval != null && !"".equals(eval)) {
            field.setAccessible(true);
            field.set(t, eval);
        }
    }

    private static void validateValue(AtomicInteger integer, String cellValue, ExcelName annotation) {
        if (annotation.required() && StringUtils.isEmpty(cellValue)) {
            throw new ParameterException("第" + integer.get() + "行" + annotation.value() + "不能为空");
        }
        if (annotation.validLen() > -1) {
            Optional.ofNullable(cellValue)
                    .map(s -> {
                        if (s.length() > annotation.validLen()) {
                            throw new ParameterException("第" + integer.get() + "行" + annotation.value() + "字符超过" + annotation.validLen() + "的最大限制");
                        }
                        return true;
                    })
                    .orElseThrow(() -> new ParameterException("第" + integer.get() + "行" + annotation.value() + "不能为空!"));
        }
    }

    private static String getCellValue(XSSFCell xssfCell) {
        CellType cellType = xssfCell.getCellType();
        if (cellType == CellType.NUMERIC) {
            //  todo 优化判断是不是日期
            String string = xssfCell.toString();
            if (string.contains("年")) {
                LocalDateTime dateTime = xssfCell.getLocalDateTimeCellValue();
                return dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
            }
            double numericCellValue = xssfCell.getNumericCellValue();
            BigDecimal bigDecimal = BigDecimal.valueOf(numericCellValue).setScale(0, BigDecimal.ROUND_HALF_UP);
            return bigDecimal.toString();
        }
        if (cellType == CellType.STRING) {
            return xssfCell.getStringCellValue();
        }
        return xssfCell.toString();
    }

    private static boolean checkRowIsEmpty(XSSFRow row) {
        if (row == null) {
            return true;
        }
        short lastCellNum = row.getLastCellNum();
        if (lastCellNum < 0) {
            return true;
        }
        boolean isRowEmpty = false;
        for (int j = 0; j < lastCellNum; j++) {
            XSSFCell cell = row.getCell(j);
            if (null == cell) {
                isRowEmpty = true;
                break;
            }
            String string = cell.toString();
            if ("".equals(string.trim())) {
                isRowEmpty = true;
            } else {
                isRowEmpty = false;
                break;
            }
        }
        return isRowEmpty;
    }


    public static <T> List<T> readAllSheetExcel(InputStream in, Class<T> clzz) {
        List<T> list = new ArrayList<>();
        XSSFWorkbook sheets = null;
        try {
            sheets = new XSSFWorkbook(in);
            for (int k = 0; k < sheets.getNumberOfSheets(); k++) {
                XSSFSheet sheetAt = sheets.getSheetAt(k);
                if (Objects.isNull(sheetAt)) {
                    continue;
                }
                convertSheet(sheetAt, clzz, list);
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

    private static boolean existsExcelAnno(String name, Field field) {
        ExcelName annotation = field.getAnnotation(ExcelName.class);
        if (null == annotation) {
            return false;
        } else {
            return annotation.value().equals(name);
        }
    }
}
