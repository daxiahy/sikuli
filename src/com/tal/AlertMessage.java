package com.tal;

import javax.swing.JOptionPane;

public class AlertMessage {
	
	public static void AlertErrorMessage(String str) {
		JOptionPane.showMessageDialog(null, str, "������ʾ", JOptionPane.ERROR_MESSAGE);
	}
	
	public static void AlertPromptMessage(String str) {
		JOptionPane.showMessageDialog(null, str, "��ʾ", JOptionPane.INFORMATION_MESSAGE);
	}
}
