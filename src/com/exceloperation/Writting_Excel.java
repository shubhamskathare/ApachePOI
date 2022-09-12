package com.exceloperation;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writting_Excel {
	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Employee_det");
		
		Object empdata[][]= {{"emp_id","emp_name","salary"},
		                 {1,"David",20000},
		                 {2,"Amith",30000},
		                 {3,"avid",40000},
		};
		int rows=empdata.length;
		int cols=empdata.length;
		
		System.out.println(rows);
		System.out.println(cols);
		
		for(int r=0;r<rows;r++){
			XSSFRow row=sheet.createRow(r);
			for(int c=0;c<cols;c++) {
		XSSFCell cell=row.createCell(c);
		
	Object value=empdata[r][c];
	
	if(value instanceof String) {
		cell.setCellValue((String)value);
		if(value instanceof Integer) {
			cell.setCellValue((Integer)value);
			if(value instanceof Boolean) {
				cell.setCellValue((Boolean)value);
	}
			}
		
		String filepath="";
		FileOutputStream outputstream=new FileOutputStream(filepath);
		workbook.write(outputstream);
		workbook.close();
		}
	}

		}
	}
}
