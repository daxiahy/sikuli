package com.tal;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	/**
	 * 向表格中写入数据
	 * @param filePath 表格路径
	 * @param nowRowNum 需要写入的行
	 * @param str 需要写入的内容
	 */
	public static void writeResult(String filePath,int nowRowNum,String str) {
		try {
			FileInputStream is = new FileInputStream(filePath);
			Workbook wb = getWorkbook(is);
			Sheet sheet1 = wb.getSheetAt(0);
			Row row = sheet1.getRow(nowRowNum);
			//Row row = sheet1.createRow(sheet1.getLastRowNum() + 1);
			//row.setHeightInPoints((short) 25);//设置行高
			// 给这一行赋值
			row.createCell(4).setCellValue(str);
			FileOutputStream os = new FileOutputStream(filePath);
			wb.write(os);
			is.close();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static Workbook getWorkbook(InputStream is) throws IOException {
		Workbook wb = null;
		if (!is.markSupported()) {
			is = new PushbackInputStream(is, 8);
		}
		if (POIFSFileSystem.hasPOIFSHeader(is)) { // Excel2003及以下版本
			wb = (Workbook) new HSSFWorkbook(is);
		} else if (POIXMLDocument.hasOOXMLHeader(is)) { // Excel2007及以上版本
			wb = new XSSFWorkbook(is);
		} else {
			throw new IllegalArgumentException("你的Excel版本目前poi无法解析！");
		}
		return wb;
	}
}
