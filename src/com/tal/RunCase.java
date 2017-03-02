package com.tal;

import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import javax.imageio.ImageIO;

import org.sikuli.script.Key;
import org.sikuli.script.Match;
import org.sikuli.script.Screen;

/**
 * 
 * 从表格中读取每一行的指令 送到AnalyzeCase中解析为指令 然后返回再执行
 */
public class RunCase {

	/**
	 * 运行程序的主要方法
	 * 
	 * @param path
	 *            是Excel的路径
	 * @throws Exception
	 */
	public static void runCase(String path) throws Exception {
		Screen s = new Screen();
		// 读取表格中的总行数
		int com_row = ReadExcel.readRowNum(path);
		// 读取应用程序的地址
		String filepath_exe = AnalyzeCase.exePath(path);
		// 读取存储图片的路径
		String img_path = AnalyzeCase.imgPath(path);
		if (com_row >= 3) {
			try {
				/** 执行应用程序 */
				Process p = Runtime.getRuntime().exec(filepath_exe);
				/**
				 * 判断指令是否大于3行 第一行：应用程序的地址 第二行：图片的路径 第三行(及以上)：命令+图片名+参数
				 */
				for (int i = 3; i <= com_row; i++) {
					/** 读取每一行的命令，返回一个list集合 */
					List<?> result = ReadExcel.readCommand(path, i);
					/** 判断是否读取成功 */
					if (!result.get(0).equals("读取Excel失败")) {
						/**
						 * 判断是哪个命令，并做相应的解析
						 */
						// 在文本框中输入字符串
						if (result.get(0).equals("type")) {
							String picture = AnalyzeCase.analyCom(path, result);
							String parameter = AnalyzeCase.getThreeColumn(path, result);
							waitEl(img_path, picture, p, i, path);
							s.type(picture, parameter);
							// 单击一个元素
						} else if (result.get(0).equals("click")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.click(picture);
							// 等待
						} else if (result.get(0).equals("wait")) {
							double time = AnalyzeCase.getWaitTime(path, result);
							s.wait(time);
							// 右键单击一个元素
						} else if (result.get(0).equals("rightClick")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.rightClick(picture);
							// 查找元素,具体作用未实现
						} else if (result.get(0).equals("find")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.find(picture);
							// 双击一个元素
						} else if (result.get(0).equals("doubleClick")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.doubleClick(picture);
							// 检查元素是否在屏幕上显示
						} else if (result.get(0).equals("exists")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.exists(picture);
							// 旋转指定的图像，功能暂时未开发，因为表格中的参数数据为确定
						} else if (result.get(0).equals("wheel")) {
							// 拖放图片
						} else if (result.get(0).equals("dragDrop")) {
							String picture1 = AnalyzeCase.analyCom(path, result);
							String picture2 = AnalyzeCase.analyCom1(path, result);
							s.dragDrop(picture1, picture2);
							// 将鼠标光标悬停在点击上方
						} else if (result.get(0).equals("hover")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.hover(picture);
							// 粘贴复制的字符串
						} else if (result.get(0).equals("paste")) {
							String picture = AnalyzeCase.analyCom(path, result);
							String parameter = AnalyzeCase.getThreeColumn(path, result);
							waitEl(img_path, picture, p, i, path);
							s.paste(picture, parameter);
							// 等待，直到该区域指定的图片消失
						} else if (result.get(0).equals("waitVanish")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.waitVanish(picture, 300);
							// 突出显示该区域一段时间
						} else if (result.get(0).equals("highlight")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.highlight(picture);
						} else if (result.get(0).equals("key")) {
							String inputKey = AnalyzeCase.getKeyMessage(path, result);
							if (inputKey.equals("enter")) {
								s.keyDown(Key.ENTER);
							}
							// 弹出提示框
						} else if (result.get(0).equals("popup")) {
							String message = AnalyzeCase.getPopupMessage(path, result);
							AlertMessage.AlertPromptMessage(message);
							// 重复单击某个按钮
						} else if (result.get(0).equals("clickNum")) {
							String picture = AnalyzeCase.analyCom(path, result);
							double time = AnalyzeCase.getClickWaitTime(path, result);
							int num = AnalyzeCase.getClickNum(path, result);
							for (int j = 0; j < num; j++) {
								s.click(picture);
								s.wait(time);
							}
							waitEl(img_path, picture, p, i, path);
						} else if (result.get(0).equals("findAll")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							for (Iterator<Match> ff = s.findAll(picture); ff.hasNext();) {
								s.click(picture);
							}

						}

					}

				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			System.out.println("命令行数不足");
			AlertMessage.AlertErrorMessage("命令行数不足！");
		}
	}

	/**
	 * 对屏幕进行拍照
	 * 
	 * @param path
	 *            图片的存储路径
	 * @throws Exception
	 *             说明：在截图后，会将图片存储到eImage目录下为err.png文件
	 */

	public static void snapShot(String path) throws Exception {

		String imaFormat = "png";
		Dimension d = Toolkit.getDefaultToolkit().getScreenSize();
		String fName = path + "eImage/";
		BufferedImage screenshot = (new Robot())
				.createScreenCapture(new Rectangle(0, 0, (int) d.getWidth(), (int) d.getHeight()));

		File folder = new File(fName);
		if (!folder.exists() && !folder.isDirectory()) {
			folder.mkdir();
		}

		String name = fName + "err." + imaFormat;
		File f = new File(name);

		// 将screenshot对象写入图像文件
		ImageIO.write(screenshot, imaFormat, f);

	}

	/**
	 * 等待元素出现，如果元素未出现则说明存在异常，然后进行屏幕截图 如果元素出现，则代表正常执行
	 * 
	 * @param scr
	 *            即为Screen
	 * @param path
	 *            图片路径，用于存储截图
	 * @param path_ima
	 *            执行过程中命令需要用到的图片
	 * @param time
	 *            需要等待的时间
	 * @param app
	 *            用于结束应用程序
	 * @param nowRow
	 *            获得当前命令所在行，为了将本行的执行结果写进去
	 * @param excelPath
	 *            表格路径，用于将结果写入到Excel中
	 * @throws Exception
	 */
	public static void waitEl(String path, String path_ima, Process app, int nowRow, String excelPath)
			throws Exception {
		// LogManager log = new LogManager();
		// Logger log = LogManager.getLogger(className);
		// DOMConfigurator.configure("log4j.xml");
		Screen s = new Screen();
		try {
			s.wait(path_ima, 0.5);
			// log.info("元素存在");
			String result = "pass";
			WriteExcel.writeResult(excelPath, nowRow, result);
		} catch (Exception e) {
			app.destroyForcibly();
			RunCase.snapShot(path);
			// log.info("元素不存在");
			String result = "fail";
			WriteExcel.writeResult(excelPath, nowRow, result);
			throw (e);
		}
	}

}
