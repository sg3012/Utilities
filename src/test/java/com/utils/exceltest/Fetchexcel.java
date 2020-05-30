package com.utils.exceltest;

import com.utils.excel.Excelops;

public class Fetchexcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Excelops read = new Excelops("D:\\Automation\\ExcelFiles\\Logindata.xlsx");
//		String data=read.getData(1,2,1); 
//		System.out.println("Data: "+data);
		String data=read.getData(1, 5, "Password"); 
		System.out.println("data: "+data);
	}

}
