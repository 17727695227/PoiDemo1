package com.example.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo13 {

	public static void main(String[] args) throws Exception{
		InputStream inp=new FileInputStream("c:\\工作簿.xls");
		POIFSFileSystem fs=new POIFSFileSystem(inp);
		Workbook wb=new HSSFWorkbook(fs);
		Sheet sheet=wb.getSheetAt(0);  // 获取第一个Sheet页
		Row row=sheet.getRow(0); // 获取第一行
		Cell cell=row.getCell(0); // 获取单元格
		if(cell==null){
			cell=row.createCell(3);
		}
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("测试单元格");
		
		FileOutputStream fileOut=new FileOutputStream("c:\\工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
