package com.zhiwei.poi;

import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class ZhiweiPoiDemo5 {

	public static void main(String[] args) throws Exception{
	Workbook workbook = new HSSFWorkbook();//定义一个工作簿
	Sheet sheet = workbook.createSheet("第一个Sheet页");//创建一个sheet页
	Row row = sheet.createRow(0);//创建一个行
	Cell cell = row.createCell(0); // 创建一个单元格  第1列
	//时间的格式化列
	CreationHelper creationHelper = workbook.getCreationHelper();
	CellStyle cellStyle=workbook.createCellStyle();//单元格样式类
	cellStyle.setDataFormat(creationHelper.createDataFormat().
			getFormat("yyy-mm-dd hh:mm:ss"));
	cell.setCellValue(new Date());//给单元格设置值
	cell.setCellStyle(cellStyle);
	
	row.createCell(1).setCellValue(1);
	row.createCell(2).setCellValue("一个字符串");
	row.createCell(3).setCellValue(HSSFCell.CELL_TYPE_NUMERIC);
	row.createCell(4).setCellValue(false);
	
	FileOutputStream fileOutputStream=new 
			FileOutputStream("f:\\工作簿.xls");
	workbook.write(fileOutputStream);
	fileOutputStream.close();
	}
}
