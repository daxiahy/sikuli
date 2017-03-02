package com.tal;

import org.sikuli.script.FindFailed;

public class Main {
		
	public static void main(String[] args) throws Exception {
		
		String paths = "D:\\sikuli.xlsx";
		
		try {
			RunCase.runCase(paths);
		} catch (FindFailed e) {
			e.printStackTrace();
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}
	
}
