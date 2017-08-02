package com.example.poi;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo4 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); // 定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个Sheet页");  // 创建第一个Sheet页
		Row row=sheet.createRow(0); // 创建一个行
		Cell cell=row.createCell(0); // 创建一个单元格  第1列
		cell.setCellValue(new Date());  // 给单元格设置值
		
		CreationHelper createHelper=wb.getCreationHelper();
		CellStyle cellStyle=wb.createCellStyle(); //单元格样式类
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyy-mm-dd hh:mm:ss"));
		cell=row.createCell(1); // 第二列
		cell.setCellValue(new Date());
		cell.setCellStyle(cellStyle);
		
		cell=row.createCell(2);  // 第三列
		cell.setCellValue(Calendar.getInstance());
		cell.setCellStyle(cellStyle);
		
		FileOutputStream fileOut=new FileOutputStream("c:\\工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
