package com.hello.common.util.excel;

import com.hello.common.util.BusinessException;
import com.hello.common.util.DateUtil;
import com.hello.common.util.ReflectUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 功能：用于2007以上版本（.xlsx）
 */
public class XssfExcelUtil extends ExcelUtil {

    public XssfExcelUtil(){}

	public XssfExcelUtil(InputStream inputStream) {
		this.inputStream = inputStream;
	}

    @Override
    public <T> List<T> readExcel(Class<T> clazz, int sheetNo, boolean hasTitle, String group) throws Exception {
        List<T> dataModels = new ArrayList<>();
        // 获取excel工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(sheetNo);
        int start = sheet.getFirstRowNum() + (hasTitle ? 1 : 0); // 如果有标题则从第二行开始
        String[] fieldNames = getClassFieldByExcelImport(clazz, group);
        for (int i = start, length = sheet.getLastRowNum(); i <= length; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 生成实例并通过反射调用setter方法
            T target = clazz.newInstance();
            for (int j = 0; j < fieldNames.length; j++) {
                String fieldName = fieldNames[j];
                if (fieldName == null || SERIALVERSIONUID.equals(fieldName)) {
                    continue; // 过滤serialVersionUID属性
                }
                XSSFCell cell = row.getCell(j); // 获取excel单元格的内容
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
        throw new BusinessException("不支持2007版本Excel导出！");
    }

	@Override
	protected Object parseValueWithType(String value, Class<?> type) {
		// 由于Excel2007的numeric类型只返回double型，所以对于类型为整型的属性，要提前对numeric字符串进行转换
		if (Byte.TYPE == type || Short.TYPE == type || Integer.TYPE == type || Long.TYPE == type) {
            value = String.valueOf((long) Double.parseDouble(value));
		}
		return super.parseValueWithType(value, type);
	}

    // ------ 待删除 ------
    /**
     * 获取单元格的内容
     *
     * @param cell
     *            单元格
     * @return 返回单元格内容
     */
    /*
	private String getCellContent(XSSFCell cell) {
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