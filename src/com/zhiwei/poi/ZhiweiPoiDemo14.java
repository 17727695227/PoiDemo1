package com.zhiwei.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class ZhiweiPoiDemo14 {

	public static void main(String[] args) throws Exception{

		Workbook workbook=new HSSFWorkbook(); // 定义一个新的工作簿
		Sheet sheet=workbook.createSheet("第一个Sheet页");// 创建第一个Sheet页
		Row row=sheet.createRow(2); // 创建一个行
		
	    Cell cell=row.createCell(2);
	    cell.setCellValue("我要换行\n 成功");
	    
	    CellStyle style=workbook.createCellStyle();
	    //设置可换行
	    style.setWrapText(true);
	    cell.setCellStyle(style);
	    
	    //调整下行的高度
	    row.setHeightInPoints(2*sheet.getDefaultRowHeightInPoints());
	    //调整单元格宽度
	    sheet.autoSizeColumn(5);
		
		FileOutputStream fileOut=new FileOutputStream("f:\\工作簿.xls");
		workbook.write(fileOut);
		fileOut.close();
	}
}
