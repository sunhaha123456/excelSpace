package com.hello.controller;

import com.hello.common.util.excel.ExcelUtil;
import com.hello.common.util.excel.HssfExcelUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

@RestController
public class ExcelDemoController {

	@RequestMapping("/excelExportA")
	public void excelExportA(HttpServletResponse response) throws Exception {
		List<ExcelReadDemo> list = new ArrayList();
		ExcelReadDemo a = new ExcelReadDemo();
		ExcelReadDemo b = new ExcelReadDemo();
		a.setHaha("哈哈");
		a.setHehe("123");
		a.setHeihei("123.12");
		a.setXixi("xixi");
		a.setDate("2019-01");
		a.setDatetime("2019-01-01 12:12:12");
		list.add(a);
		list.add(b);
		ExcelUtil excelUtil = new HssfExcelUtil();
		excelUtil.writeExcel(response, "哈哈", "呵呵", list);
	}

	@RequestMapping("/excelExportB")
	public void excelExportB(HttpServletResponse response) throws Exception {
		List<ExcelReadDemo> list = new ArrayList();
		ExcelReadDemo a = new ExcelReadDemo();
		ExcelReadDemo b = new ExcelReadDemo();
		a.setHaha("哈哈");
		a.setHehe("123");
		a.setHeihei("123.12");
		a.setXixi("xixi");
		a.setDate("2019-01");
		a.setDatetime("2019-01-01 12:12:12");
		list.add(a);
		list.add(b);
		ExcelUtil excelUtil = new HssfExcelUtil();
		excelUtil.writeExcel(response, "哈哈", "呵呵", list, "groupB", "123");
	}

	@RequestMapping("/excelImport")
	public void excelImport() throws Exception {
		File f = new File("F:/哈哈.xls");
		HssfExcelUtil excelUtil = new HssfExcelUtil(new FileInputStream(f));
		List<ExcelReadDemo> list = excelUtil.readExcel(ExcelReadDemo.class);
		System.out.println(list.toArray());
	}
}