package com.utils.exceltest;

import com.utils.excel.Excelops;

public class Writeexcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Excelops write = new Excelops("D:\\Automation\\ExcelFiles\\Logindata.xlsx");
        if(!(write.isSheetexist("Logindata")))
        {
        	 write.addsheet("Logindata"); 
        }
        
        else 
        	System.out.println("Duplicate sheet name");
	}

}
