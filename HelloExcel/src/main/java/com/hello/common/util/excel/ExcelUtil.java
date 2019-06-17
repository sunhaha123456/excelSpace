package com.hello.common.util.excel;

import com.hello.common.util.DateUtil;
import org.apache.commons.lang.ArrayUtils;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.util.IOUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 功能：excel 工具类
 * 备注：
 * 1、excel导入实体类中，导入属性必须为String类型，
 *    比如时间类型的属性，不能为Date类型，只能为String
 * 2、上传excel时，excel文档中填写日期时，对应实体类中为String：
 *    ExcelImport注解class属性配置为Date，将会转成yyyy-MM-dd格式
 *    ExcelImport注解class属性配置为Datetime，将会转成yyyy-MM-dd HH:mm:ss格式
 * 3、
 *
 *
 *
 */
public abstract class ExcelUtil {

    protected static final String SERIALVERSIONUID = "serialVersionUID";

    protected InputStream inputStream = null;

    public static ExcelUtil create(InputStream inputStream) throws IOException {
        ExcelUtil excelUtil = null;
        byte[] data = IOUtils.peekFirst8Bytes(inputStream);
        if (FileMagic.valueOf(data) == FileMagic.OLE2) {
            excelUtil = new HssfExcelUtil(inputStream);
        } else if (FileMagic.valueOf(data) == FileMagic.OOXML) {
            excelUtil = new XssfExcelUtil(inputStream);
        }
        return excelUtil;
    }

    public <T> List<T> readExcel(Class<T> clazz) throws Exception {
        return readExcel(clazz, 0, true);
    }

    public <T> List<T> readExcel(Class<T> clazz, int sheetNo, boolean hasTitle) throws Exception {
        return readExcel(clazz, sheetNo, hasTitle, "");
    }

    /**
     * 功能：抽象方法：将指定excel文件中的数据转换成数据列表，由子类实现
     * @param clazz 数据类型
     * @param sheetNo 工作表编号
     * @param hasTitle 是否带有标题
     * @param group 读取指定组名的数值，null 或 "" 则读取全部
     * @return 返回转换后的数据列表
     * @throws Exception
     */
    public abstract <T> List<T> readExcel(Class<T> clazz, int sheetNo, boolean hasTitle, String group) throws Exception;

    public <T> void writeExcel(HttpServletResponse response, String filename, String sheetName, List<T> list) throws Exception {
        writeExcel(response, filename, sheetName, list, null, null);
    }

    /**
     * 功能：抽象方法：写入数据到指定excel文件中，由子类实现
     * @param response
     * @param filename Excel表格名
     * @param sheetName Excel工作簿名
     * @param list 数据，保证 list != null && list.size > 0
     * @param groupName 读取指定组的数据，"" 或 null 表示读取全部数据
     * @param passwordOpen 表格打开密码，"" 或 "null" 或 null 表示不进行加密
     * @param <T>
     * @throws Exception
     */
    public abstract <T> void writeExcel(HttpServletResponse response, String filename, String sheetName, List<T> list, String groupName, String passwordOpen) throws Exception;

    /**
     * 功能：判断属性是否为日期类型
     * @param clazz     数据类型
     * @param fieldName 属性名
     * @return true：日期类型，false：非日期类型
     */
    protected <T> boolean isDateType(Class<T> clazz, String fieldName) {
        boolean flag = false;
        try {
            Field field = clazz.getDeclaredField(fieldName);
            ExcelImport excel = field.getAnnotation(ExcelImport.class);
            return "Date".equals(excel.clazz());
        } catch (Exception e) {
            // 把异常吞掉直接返回false
        }
        return flag;
    }

    /**
     * 功能：判断属性是否为日期时间类型
     * @param clazz     数据类型
     * @param fieldName 属性名
     * @return true：日期类型，false：非日期类型
     */
    protected <T> boolean isDateTimeType(Class<T> clazz, String fieldName) {
        boolean flag = false;
        try {
            Field field = clazz.getDeclaredField(fieldName);
            ExcelImport excel = field.getAnnotation(ExcelImport.class);
            return "Datetime".equals(excel.clazz());
        } catch (Exception e) {
            // 把异常吞掉直接返回false
        }
        return flag;
    }

    /**
     * 功能：根据类型将指定参数转换成对应的类型，供反射使用
     * @param value 指定参数
     * @param type 指定类型
     * @return 返回转换后的类型对象
     */
    protected Object parseValueWithType(String value, Class<?> type) {
        Object result = null;
        try { // 根据属性的类型将内容转换成对应的类型
            if (Boolean.TYPE == type || Boolean.class == type) {
                result = Boolean.parseBoolean(value);
            } else if (Byte.TYPE == type || Byte.class == type) {
                result = Byte.parseByte(value);
            } else if (Short.TYPE == type || Short.class == type) {
                result = Short.parseShort(value);
            } else if (Integer.TYPE == type || Integer.class == type) {
                result = Double.valueOf(value).intValue();
            } else if (Long.TYPE == type || Long.class == type) {
                result = Double.valueOf(value).longValue();
            } else if (Float.TYPE == type || Float.class == type) {
                result = Float.parseFloat(value);
            } else if (Double.TYPE == type || Double.class == type) {
                result = Double.parseDouble(value);
            } else if (Date.class == type) {
                result = DateUtil.formatStr2Date(value);
            } else {
                result = value;
            }
        } catch (Exception e) {
            // 把异常吞掉直接返回null
        }
        return result;
    }

    /**
     * 功能：获取要处理类中，指定分组的，ExcelImport注解的，类属性名称数组
     * 备注：如果类中无 ExcelImport注解，则返回全部的属性名称数组
     * @param clazz
     * @param groupName null 或 ""，则返回类ExcelImport注解下的全部属性名称数组
     * @return
     */
    protected String[] getClassFieldByExcelImport(Class clazz, String groupName) {
        Field[] fields = clazz.getDeclaredFields();
        String[] fieldNames = new String[fields.length];
        List<String> fieldAnnotations = new ArrayList<>();
        for (int i = 0; i < fields.length; i++) {
            fieldNames[i] = fields[i].getName();
            if (fields[i].isAnnotationPresent(ExcelImport.class)) {
                ExcelImport excel = fields[i].getAnnotation(ExcelImport.class);
                if (groupName == null || "".equals(groupName) || "".equals(excel.group()[0])) {
                    fieldAnnotations.add(fields[i].getName());
                } else if (ArrayUtils.contains(excel.group(), groupName)) {
                    fieldAnnotations.add(fields[i].getName());
                }
            }
        }
        if (fieldAnnotations.size() > 0) {
            fieldNames = new String[fieldAnnotations.size()];
            return fieldAnnotations.toArray(fieldNames);
        }
        return fieldNames;
    }
}