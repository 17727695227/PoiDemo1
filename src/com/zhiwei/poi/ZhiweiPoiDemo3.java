package com.zhiwei.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ZhiweiPoiDemo3 {

	public static void main(String[] args) throws Exception{
		Workbook workbook = new HSSFWorkbook();// 定义一个新的工作簿
	  	Sheet sheet = workbook.createSheet("第一个Sheet页"); // 创建第一个Sheet页
	    Row row = sheet.createRow(0);//创建一个行
	    Cell cell = row.createCell(0);//创建一个单元格  第一列
	    cell.setCellValue(1);
	    
	    row.createCell(1).setCellValue(1.2);
	    row.createCell(2).setCellValue("this is string");
	    row.createCell(3).setCellValue(false);
	    
	    FileOutputStream fileOutputStream = new FileOutputStream("f:\\用Poi搞出来的Cell.xls");
	    workbook.write(fileOutputStream);
	    fileOutputStream.close();
	}
}
