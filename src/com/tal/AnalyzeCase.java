package com.tal;

import java.util.List;
/**
 * 解析相关的sikuli命令
 * 
 */
public class AnalyzeCase {
	
	
	public static String exePath(String path) {
		/**图片地址*/
		String filepath_img = ReadExcel.readPath(path,1);
		return filepath_img;
	}
	
	/**
	 * 解析图片路径
	 * @param path表格地址
	 * @param list正在读取表格中当前行的命令的集合
	 * @return
	 */
	public static String imgPath(String path) {
		/**图片地址*/
		String filepath_img = ReadExcel.readPath(path,2);
		return filepath_img;
	}
	
	/**
	 * 解析图片路径
	 * @param path表格地址
	 * @param list正在读取表格中当前行的命令的集合
	 * @return 返回图片名及其路径
	 */
	public static String analyCom(String path,List<?> list) {
		/**图片地址*/
		String filepath_img = ReadExcel.readPath(path,2);
		/**图片名称*/
		String image_name = (String)list.get(1);
		String image = filepath_img+image_name;
		return image;
	}
	
	/**
	 * 读取type命令中需要输入的内容
	 * @param path
	 * @param list
	 * @return
	 */
	public static String getThreeColumn(String path,List<?> list) {
		String parameter = (String) list.get(2);
		return parameter;
	}
	/**
	 * 读取wait类型中的等待时间
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
	 * 解析图片路径(用于解析dragDrop命令的参数2图片)
	 * @param path表格地址
	 * @param list正在读取表格中当前行的命令的集合
	 * @return 返回图片名及其路径
	 */
	public static String analyCom1(String path,List<?> list) {
		/**图片地址*/
		String filepath_img = ReadExcel.readPath(path,2);
		/**图片吗名称*/
		String image_name = (String)list.get(2);
		String image = filepath_img+image_name;
		return image;
	}
	
	/**
	 * 读取clickNum类型中单击次数
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
	 * 读取clickNum类型中的等待时间
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
	 * 读取popup类型中的提示内容
	 * @param path
	 * @param list
	 * @return
	 */
	public static String getPopupMessage(String path,List<?> list) {
		String message = (String)list.get(1);
		return message;
	}
	
	/**
	 * 读取key
	 * @param path
	 * @param list
	 * @return
	 */
	public static String getKeyMessage(String path,List<?> list) {
		String message = (String)list.get(1);
		return message;
	}
	
}
