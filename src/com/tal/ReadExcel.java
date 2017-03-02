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
 * ˵����ReadTheCase����Ҫֻ���Ƕ�ȡ����е��������
 * @author hwl
 *
 */
public class ReadExcel {

	/** ������ */
	private int totalRows = 0;
	/** ������ */
	private int totalCells = 0;
	/** ������Ϣ */
	private String errorInfo;
	/** ������Ϣ*/
	private static String errorMessage = "��ȡExcelʧ��";

	/** ���췽�� */
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
		/** ����ļ����Ƿ�Ϊ�ջ����Ƿ���Excel��ʽ���ļ� */
		if (filePath == null || !(WDWUtil.isExcel2003(filePath) || WDWUtil.isExcel2007(filePath))) {
			errorInfo = "�ļ�������excel��ʽ";
			return false;
		}
		/** ����ļ��Ƿ���� */
		File file = new File(filePath);
		if (file == null || !file.exists()) {
			errorInfo = "�ļ�������";
			return false;
		}
		return true;
	}

	public List<List<String>> read(String filePath) {
		List<List<String>> dataLst = new ArrayList<List<String>>();
		InputStream is = null;
		try {
			/** ��֤�ļ��Ƿ�Ϸ� */
			if (!validateExcel(filePath)) {
				System.out.println(errorInfo);
				AlertMessage.AlertErrorMessage("�ļ����Ϸ���");
				return null;
			}
			/** �ж��ļ������ͣ���2003����2007 */
			boolean isExcel2003 = true;
			if (WDWUtil.isExcel2007(filePath)) {
				isExcel2003 = false;
			}
			/** ���ñ����ṩ�ĸ�������ȡ�ķ��� */
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
		/** ��������ȡ�Ľ�� */
		return dataLst;
	}

	public List<List<String>> read(InputStream inputStream, boolean isExcel2003) {
		List<List<String>> dataLst = null;
		try {
			/** ���ݰ汾ѡ�񴴽�Workbook�ķ�ʽ */
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
		/** �õ���һ��shell */
		Sheet sheet = wb.getSheetAt(0);
		/** �õ�Excel������ */
		this.totalRows = sheet.getPhysicalNumberOfRows();
		/** �õ�Excel������ */
		if (this.totalRows >= 1 && sheet.getRow(0) != null) {
			this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
		}
		/** ѭ��Excel���� */
		for (int r = 0; r < this.totalRows; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				continue;
			}
			List<String> rowLst = new ArrayList<String>();
			/** ѭ��Excel���� */
			for (int c = 0; c < this.getTotalCells(); c++) {
				Cell cell = row.getCell(c);
				String cellValue = "";
				if (null != cell) {
					/** �������ж����ݵ�����*/
					switch (cell.getCellType()) {
					case HSSFCell.CELL_TYPE_NUMERIC: // ����
						cellValue = cell.getNumericCellValue() + "";
						break;
					case HSSFCell.CELL_TYPE_STRING: // �ַ���
						cellValue = cell.getStringCellValue();
						break;
					case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
						cellValue = cell.getBooleanCellValue() + "";
						break;
					case HSSFCell.CELL_TYPE_FORMULA: // ��ʽ
						cellValue = cell.getCellFormula() + "";
						break;
					case HSSFCell.CELL_TYPE_BLANK: // ��ֵ
						cellValue = "";
						break;
					case HSSFCell.CELL_TYPE_ERROR: // ����
						cellValue = "�Ƿ��ַ�";
						break;
					default:
						cellValue = "δ֪����";
						break;
					}
				}
				rowLst.add(cellValue);
			}
			/** �����r�еĵ�c�� */
			dataLst.add(rowLst);
		}
		return dataLst;
	}

	/**
	 * ��ȡ���ÿһ�е�����
	 * @param path ���·����ַ
	 * @param num_row ���ڼ��е�����
	 * @return ����ȡ�����ͼƬ���ơ��Լ���Ҫ����Ĳ�����װ��list����
	 */
	public static List<Object> readCommand(String path, int num_row) {

		List<Object> listrow = new ArrayList<>();
		ReadExcel poi = new ReadExcel();
		List<List<String>> list = poi.read(path);
		String command = null;// �����һ�У�����
		String img_path = null;// ����ڶ���,ͼƬ��ַ��
		String parameter = null;// ��������У�����
		if (list.size() != 0) {
			List<String> cellList = list.get(num_row);
			for (int j = 0; j < 4; j++) {
				/**��ȡָ��*/
				if (j == 0) {
					command = cellList.get(j);
					listrow.add(command);
				}
				/**��ȡͼƬ��ַ*/
				if (j == 1) {
					img_path = cellList.get(j);
					listrow.add(img_path);
				}
				/**��������*/
				if (j == 2) {
					parameter = cellList.get(j);
					listrow.add(parameter);
				}
				/**��������*/
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
	 * ��ȡǰӦ�ó����ͼƬ·���ķ���
	 * 
	 * @param path
	 *            Excel��·��
	 * @param rownum
	 *            ��Ҫ��ȡ��rownum�еĵ�0��
	 * @return
	 */
	public static String readPath(String path, int rownum) {

		ReadExcel poi = new ReadExcel();
		List<List<String>> list = poi.read(path);
		String file_path = null;// �����0�еĵ�ַ
		if (list.size() != 0) {
			List<String> cellList = list.get(rownum);
			for (int j = 0; j < 1; j++) {
				// ��ȡָ��
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
	 * ��ȡ����������
	 * 
	 * @param path���·��
	 * @return ��������������
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
