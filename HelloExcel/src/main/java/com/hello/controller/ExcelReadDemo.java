package com.hello.controller;

import com.hello.common.util.excel.ExcelDataEnums;
import com.hello.common.util.excel.ExcelExport;
import com.hello.common.util.excel.ExcelImport;
import lombok.Data;

@Data
public class ExcelReadDemo {

	@ExcelExport(name = "我是haha", group = "groupB")
	@ExcelImport
	private String haha;

	@ExcelExport(name = "我是hehe")
	@ExcelImport
	private String hehe;

	@ExcelExport(name = "我是heihei", format = ExcelDataEnums.DOUBLE)
	@ExcelImport(group = "gropA")
	private String heihei;

	private String xixi;

	@ExcelExport(name = "我是date", format = ExcelDataEnums.YYYYMM)
	@ExcelImport(clazz = "Date")
	private String date;

	@ExcelExport(name = "我是datetime")
	@ExcelImport
	private String datetime;
}