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
 * �ӱ���ж�ȡÿһ�е�ָ�� �͵�AnalyzeCase�н���Ϊָ�� Ȼ�󷵻���ִ��
 */
public class RunCase {

	/**
	 * ���г������Ҫ����
	 * 
	 * @param path
	 *            ��Excel��·��
	 * @throws Exception
	 */
	public static void runCase(String path) throws Exception {
		Screen s = new Screen();
		// ��ȡ����е�������
		int com_row = ReadExcel.readRowNum(path);
		// ��ȡӦ�ó���ĵ�ַ
		String filepath_exe = AnalyzeCase.exePath(path);
		// ��ȡ�洢ͼƬ��·��
		String img_path = AnalyzeCase.imgPath(path);
		if (com_row >= 3) {
			try {
				/** ִ��Ӧ�ó��� */
				Process p = Runtime.getRuntime().exec(filepath_exe);
				/**
				 * �ж�ָ���Ƿ����3�� ��һ�У�Ӧ�ó���ĵ�ַ �ڶ��У�ͼƬ��·�� ������(������)������+ͼƬ��+����
				 */
				for (int i = 3; i <= com_row; i++) {
					/** ��ȡÿһ�е��������һ��list���� */
					List<?> result = ReadExcel.readCommand(path, i);
					/** �ж��Ƿ��ȡ�ɹ� */
					if (!result.get(0).equals("��ȡExcelʧ��")) {
						/**
						 * �ж����ĸ����������Ӧ�Ľ���
						 */
						// ���ı����������ַ���
						if (result.get(0).equals("type")) {
							String picture = AnalyzeCase.analyCom(path, result);
							String parameter = AnalyzeCase.getThreeColumn(path, result);
							waitEl(img_path, picture, p, i, path);
							s.type(picture, parameter);
							// ����һ��Ԫ��
						} else if (result.get(0).equals("click")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.click(picture);
							// �ȴ�
						} else if (result.get(0).equals("wait")) {
							double time = AnalyzeCase.getWaitTime(path, result);
							s.wait(time);
							// �Ҽ�����һ��Ԫ��
						} else if (result.get(0).equals("rightClick")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.rightClick(picture);
							// ����Ԫ��,��������δʵ��
						} else if (result.get(0).equals("find")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.find(picture);
							// ˫��һ��Ԫ��
						} else if (result.get(0).equals("doubleClick")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.doubleClick(picture);
							// ���Ԫ���Ƿ�����Ļ����ʾ
						} else if (result.get(0).equals("exists")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.exists(picture);
							// ��תָ����ͼ�񣬹�����ʱδ��������Ϊ����еĲ�������Ϊȷ��
						} else if (result.get(0).equals("wheel")) {
							// �Ϸ�ͼƬ
						} else if (result.get(0).equals("dragDrop")) {
							String picture1 = AnalyzeCase.analyCom(path, result);
							String picture2 = AnalyzeCase.analyCom1(path, result);
							s.dragDrop(picture1, picture2);
							// ���������ͣ�ڵ���Ϸ�
						} else if (result.get(0).equals("hover")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.hover(picture);
							// ճ�����Ƶ��ַ���
						} else if (result.get(0).equals("paste")) {
							String picture = AnalyzeCase.analyCom(path, result);
							String parameter = AnalyzeCase.getThreeColumn(path, result);
							waitEl(img_path, picture, p, i, path);
							s.paste(picture, parameter);
							// �ȴ���ֱ��������ָ����ͼƬ��ʧ
						} else if (result.get(0).equals("waitVanish")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.waitVanish(picture, 300);
							// ͻ����ʾ������һ��ʱ��
						} else if (result.get(0).equals("highlight")) {
							String picture = AnalyzeCase.analyCom(path, result);
							waitEl(img_path, picture, p, i, path);
							s.highlight(picture);
						} else if (result.get(0).equals("key")) {
							String inputKey = AnalyzeCase.getKeyMessage(path, result);
							if (inputKey.equals("enter")) {
								s.keyDown(Key.ENTER);
							}
							// ������ʾ��
						} else if (result.get(0).equals("popup")) {
							String message = AnalyzeCase.getPopupMessage(path, result);
							AlertMessage.AlertPromptMessage(message);
							// �ظ�����ĳ����ť
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
			System.out.println("������������");
			AlertMessage.AlertErrorMessage("�����������㣡");
		}
	}

	/**
	 * ����Ļ��������
	 * 
	 * @param path
	 *            ͼƬ�Ĵ洢·��
	 * @throws Exception
	 *             ˵�����ڽ�ͼ�󣬻ὫͼƬ�洢��eImageĿ¼��Ϊerr.png�ļ�
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

		// ��screenshot����д��ͼ���ļ�
		ImageIO.write(screenshot, imaFormat, f);

	}

	/**
	 * �ȴ�Ԫ�س��֣����Ԫ��δ������˵�������쳣��Ȼ�������Ļ��ͼ ���Ԫ�س��֣����������ִ��
	 * 
	 * @param scr
	 *            ��ΪScreen
	 * @param path
	 *            ͼƬ·�������ڴ洢��ͼ
	 * @param path_ima
	 *            ִ�й�����������Ҫ�õ���ͼƬ
	 * @param time
	 *            ��Ҫ�ȴ���ʱ��
	 * @param app
	 *            ���ڽ���Ӧ�ó���
	 * @param nowRow
	 *            ��õ�ǰ���������У�Ϊ�˽����е�ִ�н��д��ȥ
	 * @param excelPath
	 *            ���·�������ڽ����д�뵽Excel��
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
			// log.info("Ԫ�ش���");
			String result = "pass";
			WriteExcel.writeResult(excelPath, nowRow, result);
		} catch (Exception e) {
			app.destroyForcibly();
			RunCase.snapShot(path);
			// log.info("Ԫ�ز�����");
			String result = "fail";
			WriteExcel.writeResult(excelPath, nowRow, result);
			throw (e);
		}
	}

}
