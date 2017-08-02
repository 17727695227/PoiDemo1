package com.zhiwei.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


public class ZhiweiPoiDemo11 {

	public static void main(String[] args) throws Exception{

		Workbook workbook=new HSSFWorkbook(); // 定义一个新的工作簿
		Sheet sheet=workbook.createSheet("第一个Sheet页");  // 创建第一个Sheet页
		Row row=sheet.createRow(1); // 创建一个行
		
		Cell cell=row.createCell(1);
		cell.setCellValue("单元格合并测试");
	   
		sheet.addMergedRegion(new CellRangeAddress(
				1,//起始行
				2,//终止行
				1,//起始列
				2 //终止列
		));
	
		FileOutputStream fileOut=new FileOutputStream("f:\\工作簿.xls");
		workbook.write(fileOut);
		fileOut.close();
	}
}
