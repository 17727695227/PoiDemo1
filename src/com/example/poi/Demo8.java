package com.example.poi;
import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo8 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); // 定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个Sheet页");  // 创建第一个Sheet页
		Row row=sheet.createRow(2); // 创建一个行
		row.setHeightInPoints(30);
		
		createCell(wb, row, (short)0, HSSFCellStyle.ALIGN_CENTER, HSSFCellStyle.VERTICAL_BOTTOM);
		createCell(wb, row, (short)1, HSSFCellStyle.ALIGN_FILL, HSSFCellStyle.VERTICAL_CENTER);
		createCell(wb, row, (short)2, HSSFCellStyle.ALIGN_LEFT, HSSFCellStyle.VERTICAL_TOP);
		createCell(wb, row, (short)3, HSSFCellStyle.ALIGN_RIGHT, HSSFCellStyle.VERTICAL_TOP);
		
		FileOutputStream fileOut=new FileOutputStream("c:\\工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
	
	/**
	 * 创建一个单元格并为其设定指定的对其方式
	 * @param wb 工作簿
	 * @param row 行
	 * @param column  列
	 * @param halign  水平方向对其方式
	 * @param valign  垂直方向对其方式
	 */
	private static void createCell(Workbook wb,Row row,short column,short halign,short valign){
		Cell cell=row.createCell(column);  // 创建单元格
		cell.setCellValue(new HSSFRichTextString("Align It"));  // 设置值
		CellStyle cellStyle=wb.createCellStyle(); // 创建单元格样式
		cellStyle.setAlignment(halign);  // 设置单元格水平方向对其方式
		cellStyle.setVerticalAlignment(valign); // 设置单元格垂直方向对其方式
		cell.setCellStyle(cellStyle); // 设置单元格样式
	}
	
}
