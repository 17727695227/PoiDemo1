package com.zhiwei.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;



public class ZhiweiPoiDemo9 {

	public static void main(String[] args) throws Exception{

	      Workbook workbook = new HSSFWorkbook();
	      Sheet sheet=workbook.createSheet("第一个Sheet页");  // 创建第一个Sheet页
		  Row row=sheet.createRow(1); // 创建一个行	      
		  Cell cell=row.createCell(1); //创建一个单元格
		  cell.setCellValue(4);		  
		  CellStyle cellStyle = workbook.createCellStyle();
		  
		  cellStyle.setBorderBottom(CellStyle.BORDER_THIN);// 底部边框
		  cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());// 底部边框颜色
	
		  cellStyle.setBorderLeft(CellStyle.BORDER_THIN);  // 左边边框
		  cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex()); // 左边边框颜色
			
		  cellStyle.setBorderRight(CellStyle.BORDER_THIN); // 右边边框
		  cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());  // 右边边框颜色
			
		  cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED); // 上边边框
		  cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());  // 上边边框颜色
			
		  cell.setCellStyle(cellStyle);
		  FileOutputStream fileOut=new FileOutputStream("f:\\工作簿.xls");
		  workbook.write(fileOut);
		  fileOut.close();
	}
}
