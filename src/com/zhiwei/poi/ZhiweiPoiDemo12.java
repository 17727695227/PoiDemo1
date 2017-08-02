package com.zhiwei.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class ZhiweiPoiDemo12 {

	public static void main(String[] args) throws Exception{

		Workbook workbook=new HSSFWorkbook(); // 定义一个新的工作簿
		Sheet sheet=workbook.createSheet("第一个Sheet页");  // 创建第一个Sheet页
		Row row=sheet.createRow(1); // 创建一个行
		
		//创建一个字体处理类
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short)24);//设置字高
		font.setFontName("Courier New");//字体类型
        font.setItalic(true);//斜体
        font.setStrikeout(true);//斜线
        
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        
		Cell cell = row.createCell((short)1);
		cell.setCellValue("This is test of fonts");
		cell.setCellStyle(style);
        
		FileOutputStream fileOut=new FileOutputStream("f:\\工作簿.xls");
		workbook.write(fileOut);
		fileOut.close();
	}
}
