package com.zhiwei.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class ZhiweiPoiDemo13 {

	public static void main(String[] args) throws Exception{

		InputStream inputStream = new FileInputStream("f:\\工作簿.xls");
		POIFSFileSystem fileSystem = new POIFSFileSystem(inputStream);
		Workbook workbook = new HSSFWorkbook(fileSystem);
		Sheet sheet = workbook.getSheetAt(0);
		Row row = sheet.getRow(0);
		Cell cell=row.getCell(0);
		if(cell==null){
			cell=row.createCell(3);
		}
		
		cell.setCellType(Cell.CELL_TYPE_STRING);
	    cell.setCellValue("测试单元格");
	    
		FileOutputStream fileOut=new FileOutputStream("f:\\工作簿.xls");
		workbook.write(fileOut);
		fileOut.close();
	}
}
