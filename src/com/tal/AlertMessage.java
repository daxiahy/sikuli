package com.tal;

import javax.swing.JOptionPane;

public class AlertMessage {
	
	public static void AlertErrorMessage(String str) {
		JOptionPane.showMessageDialog(null, str, "错误提示", JOptionPane.ERROR_MESSAGE);
	}
	
	public static void AlertPromptMessage(String str) {
		JOptionPane.showMessageDialog(null, str, "提示", JOptionPane.INFORMATION_MESSAGE);
	}
}
