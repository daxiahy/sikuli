package com.tal;

import java.util.List;
/**
 * ������ص�sikuli����
 * 
 */
public class AnalyzeCase {
	
	
	public static String exePath(String path) {
		/**ͼƬ��ַ*/
		String filepath_img = ReadExcel.readPath(path,1);
		return filepath_img;
	}
	
	/**
	 * ����ͼƬ·��
	 * @param path����ַ
	 * @param list���ڶ�ȡ����е�ǰ�е�����ļ���
	 * @return
	 */
	public static String imgPath(String path) {
		/**ͼƬ��ַ*/
		String filepath_img = ReadExcel.readPath(path,2);
		return filepath_img;
	}
	
	/**
	 * ����ͼƬ·��
	 * @param path����ַ
	 * @param list���ڶ�ȡ����е�ǰ�е�����ļ���
	 * @return ����ͼƬ������·��
	 */
	public static String analyCom(String path,List<?> list) {
		/**ͼƬ��ַ*/
		String filepath_img = ReadExcel.readPath(path,2);
		/**ͼƬ����*/
		String image_name = (String)list.get(1);
		String image = filepath_img+image_name;
		return image;
	}
	
	/**
	 * ��ȡtype��������Ҫ���������
	 * @param path
	 * @param list
	 * @return
	 */
	public static String getThreeColumn(String path,List<?> list) {
		String parameter = (String) list.get(2);
		return parameter;
	}
	/**
	 * ��ȡwait�����еĵȴ�ʱ��
	 * @param path
	 * @param list
	 * @return
	 */
	public static double getWaitTime(String path,List<?> list) {
		String str_time = (String)list.get(1);
		double time=Double.parseDouble(str_time);
		return time;
	}
	
	/**
	 * ����ͼƬ·��(���ڽ���dragDrop����Ĳ���2ͼƬ)
	 * @param path����ַ
	 * @param list���ڶ�ȡ����е�ǰ�е�����ļ���
	 * @return ����ͼƬ������·��
	 */
	public static String analyCom1(String path,List<?> list) {
		/**ͼƬ��ַ*/
		String filepath_img = ReadExcel.readPath(path,2);
		/**ͼƬ������*/
		String image_name = (String)list.get(2);
		String image = filepath_img+image_name;
		return image;
	}
	
	/**
	 * ��ȡclickNum�����е�������
	 * @param path
	 * @param list
	 * @return
	 */
	public static int getClickNum(String path,List<?> list) {
		String str_num = (String)list.get(3);
		int num=Integer.parseInt(str_num);
		return num;
	}
	
	/**
	 * ��ȡclickNum�����еĵȴ�ʱ��
	 * @param path
	 * @param list
	 * @return
	 */
	public static double getClickWaitTime(String path,List<?> list) {
		String str_time = (String)list.get(2);
		double time=Double.parseDouble(str_time);
		return time;
	}
	
	/**
	 * ��ȡpopup�����е���ʾ����
	 * @param path
	 * @param list
	 * @return
	 */
	public static String getPopupMessage(String path,List<?> list) {
		String message = (String)list.get(1);
		return message;
	}
	
	/**
	 * ��ȡkey
	 * @param path
	 * @param list
	 * @return
	 */
	public static String getKeyMessage(String path,List<?> list) {
		String message = (String)list.get(1);
		return message;
	}
	
}
