package com.zhiwei.poi;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;



public class ZhiweiPoiDemo6 {

	public static void main(String[] args) throws Exception{
		
		InputStream isInputStream = new FileInputStream("f:\\工作簿.xls");
		POIFSFileSystem fsFileSystem= new POIFSFileSystem(isInputStream);
		HSSFWorkbook wbHssfWorkbook=new HSSFWorkbook(fsFileSystem);
		HSSFSheet hasSheet=wbHssfWorkbook.getSheetAt(0);//获取一个Sheet页
		if(hasSheet==null)
		{
			return;
		}
		//遍历行row
		for (int i = 0; i <=hasSheet.getLastRowNum(); i++) {
			HSSFRow hssfRow=hasSheet.getRow(i);
			if(hssfRow==null){
				continue;
			}
			//遍历列cell
			for(int cellNum=0;cellNum<=hssfRow.getLastCellNum();cellNum++)
			{
				HSSFCell hssfCell=hssfRow.getCell(cellNum);
				if(hssfCell==null){
					continue;
				}
				System.out.print(" "+getValue(hssfCell));
			}
			System.out.println();
		}
		
		
		
	}
    //判断是什么类型
	private static String getValue(HSSFCell hssfCell) {

		if(hssfCell.getCellType()==HSSFCell.CELL_TYPE_BOOLEAN){
			return String.valueOf(hssfCell.getBooleanCellValue());
		}else if (hssfCell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC) {
			return String.valueOf(hssfCell.getNumericCellValue());
		}else if (hssfCell.getCellType()==HSSFCell.CELL_TYPE_STRING) {
			return String.valueOf(hssfCell.getStringCellValue());
		}
		return String.valueOf(hssfCell.getDateCellValue());
	}
}
