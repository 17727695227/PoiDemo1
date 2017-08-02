package com.zhiwei.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class ZhiweiPoiDemo1 {

	public static void main(String[] args) throws Exception{
		Workbook workbook = new HSSFWorkbook();//定义一个新的工作簿
		FileOutputStream fileOutputStream = new 
				FileOutputStream("f:\\用Poi搞出来的工作簿.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
	}
}
