package com.zhiwei.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class ZhiweiPoiDemo2 {

	public static void main(String[] args) throws Exception{
		Workbook workbook = new HSSFWorkbook();
		workbook.createSheet("第一个Sheet页");
		workbook.createSheet("第二个Sheet页");
		FileOutputStream fileOutputStream=new FileOutputStream("f:\\poi-sheet页.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();

	}
}
