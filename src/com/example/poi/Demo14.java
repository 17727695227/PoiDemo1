package com.example.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo14 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); // 定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个Sheet页");  // 创建第一个Sheet页
		Row row=sheet.createRow(2); // 创建一个行
		Cell cell=row.createCell(2);
		cell.setCellValue("我要换行 \n 成功了吗？");
		
		CellStyle cs=wb.createCellStyle();
		// 设置可以换行
		cs.setWrapText(true);
		cell.setCellStyle(cs);
		
		// 调整下行的高度
		row.setHeightInPoints(2*sheet.getDefaultRowHeightInPoints());
		// 调整单元格宽度
		sheet.autoSizeColumn(2);
		
		FileOutputStream fileOut=new FileOutputStream("c:\\工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
