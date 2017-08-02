package com.zhiwei.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


public class ZhiweiPoiDemo15 {

	public static void main(String[] args) throws Exception{

		Workbook workbook=new HSSFWorkbook(); // 定义一个新的工作簿
		Sheet sheet=workbook.createSheet("第一个Sheet页");  // 创建第一个Sheet页
	
		CellStyle style;
		DataFormat format=workbook.createDataFormat();
		Row row;
		Cell cell;
		short rowNum=0;
		short cellNum=0;
		
		row=sheet.createRow(rowNum++);
		cell=row.createCell(cellNum++);
		cell.setCellValue(1111.25);
		
		style=workbook.createCellStyle();
		style.setDataFormat(format.getFormat("0.0"));
		cell.setCellStyle(style);
		
		row=sheet.createRow(rowNum++);
		cell=row.createCell(cellNum);
		cell.setCellValue(1111111.25);
		style=workbook.createCellStyle();
		style.setDataFormat(format.getFormat("#,##0.000"));
		cell.setCellStyle(style);
	
		FileOutputStream fileOut=new FileOutputStream("f:\\工作簿.xls");
		workbook.write(fileOut);
		fileOut.close();
	}
}
