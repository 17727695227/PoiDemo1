package com.example.poi;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo1 {

	public static void main(String[] args) throws Exception {
		Workbook wb=new HSSFWorkbook(); //定义一个新的工作簿
		FileOutputStream fileOut=new FileOutputStream("f:\\用Poi搞出来的工作簿2.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
