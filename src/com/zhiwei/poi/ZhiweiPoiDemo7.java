package com.zhiwei.poi;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


public class ZhiweiPoiDemo7 {

	public static void main(String[] args) throws Exception{
	      InputStream isInputStream=new FileInputStream("f:\\工作簿.xls");
	      POIFSFileSystem fsFileSystem = new POIFSFileSystem(isInputStream);
	      HSSFWorkbook workbook = new HSSFWorkbook(fsFileSystem);
	      
	      ExcelExtractor excelExtractor = new ExcelExtractor(workbook);
	      excelExtractor.setIncludeSheetNames(false);//不打印“第一个Sheet页“
	}
}
