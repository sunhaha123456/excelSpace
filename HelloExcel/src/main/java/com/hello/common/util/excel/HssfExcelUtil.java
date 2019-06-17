package com.hello.common.util.excel;

import com.hello.common.util.*;
import com.hello.common.util.DateUtil;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.ArrayUtils;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import javax.servlet.http.HttpServletResponse;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * 功能：用于97-2003之间（.xls）
 */
public class HssfExcelUtil extends ExcelUtil {

    private final static String FILE_TYPE = ".xls";

    private final static String CELL_BLOCK_PASSWORD = "123456111";

    private final static String MESSAGE = "请选择或输入有效的选项！";

    private final static Integer MAX_ROW = 50000;

    private static HSSFCellStyle contentStyleDefault;

    private static HSSFCellStyle contentStyleDate;

    private static HSSFCellStyle contentStyleDouble;

    public HssfExcelUtil() {}

    public HssfExcelUtil(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    @Override
    public <T> List<T> readExcel(Class<T> clazz, int sheetNo, boolean hasTitle, String group) throws Exception {
        List<T> dataModels = new ArrayList<>();
        // 获取excel工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = workbook.getSheetAt(sheetNo);
        int start = sheet.getFirstRowNum() + (hasTitle ? 1 : 0); // 如果有标题则从第二行开始
        String[] fieldNames = getClassFieldByExcelImport(clazz, group);
        for (int x = start; x <= sheet.getLastRowNum(); x++) {
            HSSFRow row = sheet.getRow(x);
            if (row == null) {
                continue;
            }
            // 生成实例并通过反射调用setter方法
            T target = clazz.newInstance();
            for (int y = 0; y < fieldNames.length; y++) {
                String fieldName = fieldNames[y];
                if (fieldName == null || SERIALVERSIONUID.equals(fieldName)) {
                    continue; // 过滤serialVersionUID属性
                }
                HSSFCell cell = row.getCell(y); // 获取excel单元格的内容
                if (cell == null || cell.toString().length() == 0) {
                    continue;
                }
                if (isDateType(clazz, fieldName) || isDateTimeType(clazz, fieldName)) {
                    boolean dateFlag = false;
                    try {
                        dateFlag = HSSFDateUtil.isCellDateFormatted(cell);
                    } catch (Exception e) {
                        throw new BusinessException("导入Excel单元格格式非时间类型！");
                    }
                    if (!dateFlag) {
                        throw new BusinessException("导入Excel单元格格式被认为是非时间类型！");
                    }
                    if (dateFlag) {
                        // 当导入实体类配置了Date 或 Datetime，并且导入excel单元格格式是时间类型时
                        double d = cell.getNumericCellValue();
                        Date date = HSSFDateUtil.getJavaDate(d);
                        String dataStr = "";
                        if (isDateType(clazz, fieldName)) {
                            dataStr = DateUtil.format(date, DateUtil.DATEFORMAT);
                        } else {
                            dataStr = DateUtil.format(date, DateUtil.TIMEFORMAT);
                        }
                        ReflectUtil.invokeSetter(target, fieldName, dataStr);
                        continue;
                    }
                } else {
                    cell.setCellType(CellType.STRING);
                    String content = cell.getStringCellValue();
                    Field field = clazz.getDeclaredField(fieldName);
                    ReflectUtil.invokeSetter(target, fieldName, parseValueWithType(content, field.getType()));
                }
            }
            dataModels.add(target);
        }
        return dataModels;
    }

    @Override
    public <T> void writeExcel(HttpServletResponse response, String filename, String sheetName, List<T> list, String groupName, String passwordOpen) throws Exception {
        // reponse init
        response.setContentType("octets/stream");
        response.addHeader("Content-Type", "octets/stream; charset=utf-8");
        filename = new String(filename.getBytes("UTF-8"), "iso8859-1");
        response.addHeader("Content-Disposition", "attachment;filename=" + filename + FILE_TYPE);
        OutputStream outputStream = response.getOutputStream();
        // 声明一个工作薄
        contentStyleDefault = null;
        contentStyleDate = null;
        contentStyleDouble = null;
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<String> headers = new ArrayList<>(); // 字段中文名
        List<FieldObject> fields = new ArrayList<>(); // 字段名
        boolean isLock = getExcelExportData(list.get(0).getClass(), headers, fields, groupName);
        if (headers.size() == 0) {
            throw new Exception("该类没有Excel导出注解请检查!");
        }
        setSheet(workbook, sheetName, headers, list, fields, isLock);
        if (StringUtil.isEmpty(passwordOpen)) {
            Biff8EncryptionKey.setCurrentUserPassword(null);
        } else {
            Biff8EncryptionKey.setCurrentUserPassword(passwordOpen);
        }
        try {
            workbook.write(outputStream);
            outputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            outputStream.close();
        }
    }

    /**
     * 功能：获取要到处数据的头部标题，以及标题下的数据
     * @param c
     * @param headers
     * @param fields
     * @param groupName
     * @return 导出的Excel是否需要进行不可修改的锁定
     */
    private boolean getExcelExportData(Class<?> c, List<String> headers, List<FieldObject> fields, String groupName) {
        boolean isLock = false;
        Field[] allFields = c.getDeclaredFields();
        int index = 0;
        for (Field field : allFields) {
            if (!field.isAnnotationPresent(ExcelExport.class)) {
                continue;
            }
            ExcelExport excel = field.getAnnotation(ExcelExport.class);
            FieldObject fieldObject;
            if (groupName == null || "".equals(groupName) || "".equals(excel.group()[0]) || ArrayUtils.contains(excel.group(), groupName)) {
                fieldObject = new FieldObject();
                if (excel.lockBoolean()) {
                    isLock = true;
                }
                fieldObject.setIndex(index++);
                fieldObject.setName(field.getName());
                fieldObject.setFormat(excel.format());
                fieldObject.setLockBoolean(excel.lockBoolean());
                fieldObject.setConstantSelectList(Arrays.asList(excel.constantSelectList()));
                fieldObject.setReturnSelectDataClass(excel.returnSelectDataClass());
                fieldObject.setReturnSelectMapClass(excel.returnSelectMapClass());
                fieldObject.setParentName(excel.parentName());
                fieldObject.setPointOut(excel.pointOut());
                fields.add(fieldObject);
                headers.add(excel.name());
            }
        }
        return isLock;
    }

    /**
     * 功能：表格赋值，对标题以及标题下的内容进行赋值
     * @param workbook
     * @param sheetName
     * @param headers
     * @param list
     * @param fields
     * @param isLock
     * @param <T>
     */
    private <T> void setSheet(HSSFWorkbook workbook, String sheetName, List<String> headers, List<T> list, List<FieldObject> fields, boolean isLock) throws Exception {
        // 生成一个表格
        HSSFSheet sheet = workbook.createSheet(sheetName);
        // 表头样式
        HSSFCellStyle headerStyle = getHeadStyle(workbook);
        // 产生表格标题行
        HSSFRow row = sheet.createRow(0);
        if (isLock) {
            sheet.protectSheet(CELL_BLOCK_PASSWORD);
        }
        HSSFRichTextString text;
        Map<String, List<String>> parentMap = new HashMap<>();
        HSSFDataFormat format = workbook.createDataFormat();
        for (int i = 0; i < headers.size(); i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(headerStyle);
            text = new HSSFRichTextString(headers.get(i));
            cell.setCellValue(text);
            //设置列宽度自适应
            sheet.setColumnWidth(i, headers.get(i).getBytes().length * 256);
            sheet.setDefaultColumnStyle(i, getCellStyle(workbook, sheet, format, fields.get(i), parentMap));
        }
        // 循环赋值
        int rowCount = 1;
        for (int i = 0; i < list.size(); i++) {
            // 正文内容样式
            row = sheet.createRow(rowCount++);
            setRowValue(row, list.get(i), fields);
        }
        //目前只展示第一个sheet
        int sheetCount = workbook.getNumberOfSheets();
        for (int i = 1; i < sheetCount; i++) {
            workbook.setSheetHidden(i, true);
            HSSFSheet hiddenSheet = workbook.getSheetAt(i);
            hiddenSheet.protectSheet(CELL_BLOCK_PASSWORD);
        }
    }

    // 标题样式
    private HSSFCellStyle getHeadStyle(HSSFWorkbook workbook) {
        HSSFCellStyle headerStyle = getCellBaseStyle(workbook);
        headerStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        headerStyle.setFont(font);
        return headerStyle;
    }

    //基本样式
    private HSSFCellStyle getCellBaseStyle(HSSFWorkbook workbook) {
        HSSFCellStyle contentStyle = workbook.createCellStyle();
        contentStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        contentStyle.setBorderBottom(BorderStyle.THIN);
        contentStyle.setBorderLeft(BorderStyle.THIN);
        contentStyle.setBorderRight(BorderStyle.THIN);
        contentStyle.setBorderTop(BorderStyle.THIN);
        contentStyle.setAlignment(HorizontalAlignment.CENTER);
        return contentStyle;
    }

    // 填写sheet的每行的值
    private void setRowValue(HSSFRow row, Object obj, List<FieldObject> fields) {
        Class<?> c = obj.getClass();
        Object value;
        PropertyDescriptor pd;
        for (int i = 0; i < fields.size(); i++) {
            try {
                HSSFCell cell = row.createCell(i);
                pd = new PropertyDescriptor(fields.get(i).getName(), c);
                Method getMethod = pd.getReadMethod();// 获得get方法
                value = getMethod.invoke(obj);
                if (value instanceof Double) {
                    cell.setCellValue((Double) value);
                } else if (value instanceof Date) {
                    cell.setCellValue((Date) value);
                } else {
                    cell.setCellValue(value == null ? "" : value.toString());
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    //内容样式
    private HSSFCellStyle getContentStyle(HSSFWorkbook workbook) {
        HSSFCellStyle contentStyle = getCellBaseStyle(workbook);
        contentStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        HSSFFont font = workbook.createFont();
        font.setBold(false);
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        contentStyle.setFont(font);
        return contentStyle;
    }

    private CellStyle getCellStyle(HSSFWorkbook workbook, HSSFSheet sheet, HSSFDataFormat format, FieldObject fileObject, Map<String, List<String>> parentMap) throws Exception {
        ExcelDataEnums enums = fileObject.getFormat();
        switch (enums) {
            case YYYYMM: {
                if (contentStyleDate == null) {
                    contentStyleDate = getContentStyle(workbook);
                    contentStyleDate.setDataFormat(format.getFormat("yyyy-mm"));
                }
                return contentStyleDate;
            }
            case DOUBLE: {
                if (contentStyleDouble == null) {
                    contentStyleDouble = getContentStyle(workbook);
                    contentStyleDouble.setDataFormat(format.getFormat("0.00"));
                }
                return contentStyleDouble;
            }
            case CONSTANT_SELECT: {
                if (contentStyleDefault == null) {
                    contentStyleDefault = getContentStyle(workbook);
                }
                if (CollectionUtils.isEmpty(fileObject.getConstantSelectList())) {
                    throw new Exception("下拉列表不能为空!");
                }
                writeSelects(workbook, fileObject.getName(), fileObject.getConstantSelectList());
                initSheetNameMapping(sheet, fileObject);
                return contentStyleDefault;
            }
            case SINGLE_SELECT:
            case UNION_PARENT_SELECT: {
                if (contentStyleDefault == null) {
                    contentStyleDefault = getContentStyle(workbook);
                }
                if (fileObject.getReturnSelectDataClass().isInterface()) {
                    throw new Exception("下拉列表类未实现!");
                }
                String className = fileObject.getReturnSelectDataClass().getSimpleName();
                ExcelSelectInterface bean = (ExcelSelectInterface) ChanelContext.getBean(className.substring(0, 1).toLowerCase() + className.substring(1));
                List<String> selects = bean.returnSelectData();
                if (selects == null || selects.size() == 0) {
                    throw new Exception("选项中没有数据!");
                }
                writeSelects(workbook, fileObject.getName(), selects);
                initSheetNameMapping(sheet, fileObject);
                if (enums.equals(ExcelDataEnums.UNION_PARENT_SELECT)) {
                    parentMap.put(fileObject.getName(), selects);
                }
                return contentStyleDefault;
            }
            case UNION_CHILD_SELECT: {
                if (contentStyleDefault == null) {
                    contentStyleDefault = getContentStyle(workbook);
                }
                String className = fileObject.getReturnSelectMapClass().getSimpleName();
                ExcelSelectMapInterface bean = (ExcelSelectMapInterface) ChanelContext.getBean(className.substring(0, 1).toLowerCase() + className.substring(1));
                Map<String, List<String>> map = bean.returnSelectMapData();
                if (map == null || map.size() == 0) {
                    throw new Exception("关联选项列表中没有数据!");
                }
                fileObject.setMap(map);
                writeSelectsMap(workbook, fileObject, parentMap);
            }
            default: {
                if (contentStyleDefault == null) {
                    contentStyleDefault = getContentStyle(workbook);
                }
                return contentStyleDefault;
            }
        }
    }

    private void writeSelects(HSSFWorkbook workbook, String sheetName, List<String> selects) {
        HSSFSheet sheet = workbook.createSheet(sheetName);
        for (int i = 0; i < selects.size(); i++) {
            HSSFRow row = sheet.createRow(i);
            HSSFCell cell1 = row.createCell(0);
            cell1.setCellValue(selects.get(i));
        }
        Name name = workbook.createName();
        name.setNameName(sheetName);
        name.setRefersToFormula(sheet.getSheetName() + "!$A$1:$A$" + selects.size());
    }

    private void writeSelectsMap(HSSFWorkbook workbook, FieldObject fileObject, Map<String, List<String>> parentMap) throws Exception {
        String parentName = fileObject.getParentName();
        List<String> parentList = parentMap.get(parentName);
        if (CollectionUtils.isEmpty(parentList)) {
            throw new Exception("联动parent list未找到！");
        }
        HSSFSheet wsSheet = workbook.getSheet(parentName);
        for (int i = 0; i < parentList.size(); i++) {
            int referColNum = i + 1;
            String parent = parentList.get(i);
            int rowCount = wsSheet.getLastRowNum();
            Map<String, List<String>> map = fileObject.getMap();
            List<String> sub = map.get(parent);
            if (!CollectionUtils.isEmpty(sub)) {
                for (int j = 0; j < sub.size(); j++) {
                    if (j <= rowCount) { //前面创建过的行，直接获取行，创建列
                        wsSheet.getRow(j).createCell(referColNum).setCellValue(sub.get(j)); //设置对应单元格的值
                    } else { //未创建过的行，直接创建行、创建列
                        wsSheet.setColumnWidth(j, 4000); //设置每列的列宽
                        //创建行、创建列
                        wsSheet.createRow(j).createCell(referColNum).setCellValue(sub.get(j)); //设置对应单元格的值
                    }
                }
                Name name = workbook.createName();
                name.setNameName(parent);
                String referColName = getColumnName(referColNum);
                String formula = parentName + "!$" + referColName + "$1:$" + referColName + "$" + sub.size();
                name.setRefersToFormula(formula);
            }
        }
    }

    private void initSheetNameMapping(HSSFSheet mainSheet, FieldObject fileObject) {
        DataValidation warehouseValidation = getDataValidationByFormula(fileObject);
        // 主sheet添加验证数据
        mainSheet.addValidationData(warehouseValidation);
    }

    private DataValidation getDataValidationByFormula(FieldObject fileObject) {
        // 加载下拉列表内容
        DVConstraint constraint;
        if (fileObject.getName().equals("term")) {
            constraint = DVConstraint.createFormulaListConstraint("INDIRECT($A1)");
        } else {
            constraint = DVConstraint.createFormulaListConstraint(fileObject.getName());
        }

        // 设置数据有效性加载在哪个单元格上。
        // 四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList(1, MAX_ROW, fileObject.getIndex(), fileObject.getIndex());
        // 数据有效性对象
        DataValidation dataValidationList = new HSSFDataValidation(regions, constraint);
        dataValidationList.createErrorBox("Error", MESSAGE);
        if (StringUtil.isEmpty(fileObject.getPointOut())) {
            dataValidationList.createPromptBox("提示", MESSAGE);
        } else {
            dataValidationList.createPromptBox("提示", fileObject.getPointOut());
        }
        return dataValidationList;
    }

    /**
     * 根据数据值确定单元格位置（比如：0-A, 27-AB）
     *
     * @param index
     * @return
     */
    private String getColumnName(int index) {
        StringBuilder s = new StringBuilder();
        while (index >= 26) {
            s.insert(0, (char) ('A' + index % 26));
            index = index / 26 - 1;
        }
        s.insert(0, (char) ('A' + index));
        return s.toString();
    }

    // --- 待删除 start ---
    /*
    private <T> boolean checkIsAddList(T target, List<String> andNullFiled) throws Exception {
        if (andNullFiled.size() > 0) {
            for (String fieldName : andNullFiled) {
                Object o = ReflectUtil.invokeGetter(target, fieldName);
                if (o != null) {
                    return true;
                }
            }
            return false;
        }
        return true;
    }
    */

    // 功能：获取单元格的内容
    // @param cell 单元格
    // @return 返回单元格内容
    /*
    private String getCellContent(HSSFCell cell) {
        Object obj = null;
        switch (cell.getCellTypeEnum()) {
            case NUMERIC : // 数字
                DecimalFormat df = new DecimalFormat("0");
                obj = df.format(cell.getNumericCellValue());
                break;
            case BOOLEAN : // 布尔
                obj = cell.getBooleanCellValue();
                break;
            case FORMULA : // 公式
                obj = cell.getCellFormula() ;
                break;
            case STRING : // 字符串
                obj = cell.getStringCellValue();
                break;
            case BLANK : // 空值
            case ERROR : // 故障
            default :
                break;
        }
        return obj + "";
    }
    */
}