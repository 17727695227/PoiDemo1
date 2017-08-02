package com.example.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo2 {

	public static void main(String[] args) throws Exception {
		
		Workbook wb=new HSSFWorkbook(); // 定义一个新的工作簿
		wb.createSheet("第一个Sheet页");  // 创建第一个Sheet页
		wb.createSheet("第二个Sheet页");  // 创建第二个Sheet页
		FileOutputStream fileOut=new FileOutputStream("f:\\用Poi搞出来的Sheet页.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
