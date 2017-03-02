package com.tal;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 说明：ReadTheCase的主要只用是读取表格中的相关内容
 * @author hwl
 *
 */
public class ReadExcel {

	/** 总行数 */
	private int totalRows = 0;
	/** 总列数 */
	private int totalCells = 0;
	/** 错误信息 */
	private String errorInfo;
	/** 错误信息*/
	private static String errorMessage = "读取Excel失败";

	/** 构造方法 */
	public ReadExcel() {
	}

	public int getTotalRows() {
		return totalRows;
	}

	public int getTotalCells() {
		return totalCells;
	}

	public String getErrorInfo() {
		return errorInfo;
	}

	public boolean validateExcel(String filePath) {
		/** 检查文件名是否为空或者是否是Excel格式的文件 */
		if (filePath == null || !(WDWUtil.isExcel2003(filePath) || WDWUtil.isExcel2007(filePath))) {
			errorInfo = "文件名不是excel格式";
			return false;
		}
		/** 检查文件是否存在 */
		File file = new File(filePath);
		if (file == null || !file.exists()) {
			errorInfo = "文件不存在";
			return false;
		}
		return true;
	}

	public List<List<String>> read(String filePath) {
		List<List<String>> dataLst = new ArrayList<List<String>>();
		InputStream is = null;
		try {
			/** 验证文件是否合法 */
			if (!validateExcel(filePath)) {
				System.out.println(errorInfo);
				AlertMessage.AlertErrorMessage("文件不合法！");
				return null;
			}
			/** 判断文件的类型，是2003还是2007 */
			boolean isExcel2003 = true;
			if (WDWUtil.isExcel2007(filePath)) {
				isExcel2003 = false;
			}
			/** 调用本类提供的根据流读取的方法 */
			File file = new File(filePath);
			is = new FileInputStream(file);
			dataLst = read(is, isExcel2003);
			is.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					is = null;
					e.printStackTrace();
				}
			}
		}
		/** 返回最后读取的结果 */
		return dataLst;
	}

	public List<List<String>> read(InputStream inputStream, boolean isExcel2003) {
		List<List<String>> dataLst = null;
		try {
			/** 根据版本选择创建Workbook的方式 */
			Workbook wb = null;
			if (isExcel2003) {
				wb = new HSSFWorkbook(inputStream);
			} else {
				wb = new XSSFWorkbook(inputStream);
			}
			dataLst = read(wb);
		} catch (IOException e) {

			e.printStackTrace();
		}
		return dataLst;
	}

	private List<List<String>> read(Workbook wb) {
		List<List<String>> dataLst = new ArrayList<List<String>>();
		/** 得到第一个shell */
		Sheet sheet = wb.getSheetAt(0);
		/** 得到Excel的行数 */
		this.totalRows = sheet.getPhysicalNumberOfRows();
		/** 得到Excel的列数 */
		if (this.totalRows >= 1 && sheet.getRow(0) != null) {
			this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
		}
		/** 循环Excel的行 */
		for (int r = 0; r < this.totalRows; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				continue;
			}
			List<String> rowLst = new ArrayList<String>();
			/** 循环Excel的列 */
			for (int c = 0; c < this.getTotalCells(); c++) {
				Cell cell = row.getCell(c);
				String cellValue = "";
				if (null != cell) {
					/** 以下是判断数据的类型*/
					switch (cell.getCellType()) {
					case HSSFCell.CELL_TYPE_NUMERIC: // 数字
						cellValue = cell.getNumericCellValue() + "";
						break;
					case HSSFCell.CELL_TYPE_STRING: // 字符串
						cellValue = cell.getStringCellValue();
						break;
					case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
						cellValue = cell.getBooleanCellValue() + "";
						break;
					case HSSFCell.CELL_TYPE_FORMULA: // 公式
						cellValue = cell.getCellFormula() + "";
						break;
					case HSSFCell.CELL_TYPE_BLANK: // 空值
						cellValue = "";
						break;
					case HSSFCell.CELL_TYPE_ERROR: // 故障
						cellValue = "非法字符";
						break;
					default:
						cellValue = "未知类型";
						break;
					}
				}
				rowLst.add(cellValue);
			}
			/** 保存第r行的第c列 */
			dataLst.add(rowLst);
		}
		return dataLst;
	}

	/**
	 * 读取表格每一行的命令
	 * @param path 表格路径地址
	 * @param num_row 读第几行的内容
	 * @return 将读取的命令、图片名称、以及需要输入的参数封装成list返回
	 */
	public static List<Object> readCommand(String path, int num_row) {

		List<Object> listrow = new ArrayList<>();
		ReadExcel poi = new ReadExcel();
		List<List<String>> list = poi.read(path);
		String command = null;// 定义第一列，命令
		String img_path = null;// 定义第二列,图片地址列
		String parameter = null;// 定义第三列，参数
		if (list.size() != 0) {
			List<String> cellList = list.get(num_row);
			for (int j = 0; j < 4; j++) {
				/**提取指令*/
				if (j == 0) {
					command = cellList.get(j);
					listrow.add(command);
				}
				/**提取图片地址*/
				if (j == 1) {
					img_path = cellList.get(j);
					listrow.add(img_path);
				}
				/**输入内容*/
				if (j == 2) {
					parameter = cellList.get(j);
					listrow.add(parameter);
				}
				/**输入内容*/
				if (j == 3) {
					parameter = cellList.get(j);
					listrow.add(parameter);
				}
			}
			return listrow;
		} else {
			listrow.add(errorMessage);
			return listrow;
		}

	}

	/**
	 * 读取前应用程序和图片路径的方法
	 * 
	 * @param path
	 *            Excel的路径
	 * @param rownum
	 *            需要读取第rownum行的第0列
	 * @return
	 */
	public static String readPath(String path, int rownum) {

		ReadExcel poi = new ReadExcel();
		List<List<String>> list = poi.read(path);
		String file_path = null;// 定义第0列的地址
		if (list.size() != 0) {
			List<String> cellList = list.get(rownum);
			for (int j = 0; j < 1; j++) {
				// 提取指令
				if (j == 0) {
					file_path = cellList.get(j);
				}
			}
			return file_path;
		} else {
			return errorMessage;
		}
	}

	/**
	 * 读取表格的总行数
	 * 
	 * @param path表格路径
	 * @return 返回命令总行数
	 */
	public static int readRowNum(String path) {
		ReadExcel poi = new ReadExcel();
		List<List<String>> list = poi.read(path);
		int numrow = list.size() - 1;
		if (list.size() != 0) {
			return numrow;
		} else {
			return 0;
		}
	}

}

class WDWUtil {

	public static boolean isExcel2003(String filePath) {
		return filePath.matches("^.+\\.(?i)(xls)$");
	}

	public static boolean isExcel2007(String filePath) {
		return filePath.matches("^.+\\.(?i)(xlsx)$");
	}

}
