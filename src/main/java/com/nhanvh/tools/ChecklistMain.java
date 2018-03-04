package com.nhanvh.tools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ChecklistMain {

	@SuppressWarnings({ "static-access", "resource", "deprecation" })
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		/*
		 * Created By: NhanVH 
		 * Description:
		 */
		CheckCell cc = new CheckCell();
		String input;
		Scanner user_input = new Scanner(System.in);
		System.out.println("Enter your file to check: ");
		input = user_input.next();
		String output;
		System.out.println("Enter your file to show result: ");
		output = user_input.next();
		
		/*
		 * Created By: NhanVH 
		 * Description: Open File Input, set Cell Style
		 */
		FileInputStream fis = new FileInputStream(new File(input));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		// Get first sheet of file input
		XSSFSheet ws = wb.getSheetAt(0);
		XSSFCellStyle cellStyle = wb.createCellStyle();
		// Fill cell by red color
		cellStyle.setFillPattern(XSSFCellStyle.LEAST_DOTS);
		cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		cellStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
		/*
		 * Created By: NhanVH
		 * Description: Open File Output, set Cell Style
		 */
		FileInputStream fis1 = new FileInputStream(new File(output));
		XSSFWorkbook wb1 = new XSSFWorkbook(fis1);
		// Get first sheet of file output
		XSSFSheet ws1 = wb1.getSheetAt(0);
		XSSFCellStyle cellStyle1 = wb1.createCellStyle();
		// Fill cell by red color
		cellStyle1.setFillPattern(XSSFCellStyle.LEAST_DOTS);
		cellStyle1.setFillForegroundColor(IndexedColors.RED.getIndex());
		cellStyle1.setFillBackgroundColor(IndexedColors.RED.getIndex());
		
		// Get & Check data on first sheet of input file
		XSSFCell c1 = ws.getRow(5).getCell(5);
		String cl1;
		if (c1 == null) {
			cl1 = "";
		} else {
			cl1 = cc.cellToString(c1);
		}
		XSSFCell c2 = ws.getRow(5).getCell(11);
		String cl2;
		if (c2 == null) {
			cl2 = "";
		} else {
			cl2 = cc.cellToString(c2);
		}
		XSSFCell c3 = ws.getRow(7).getCell(5);
		String cl3;
		if (c3 == null) {
			cl3 = "";
		} else {
			cl3 = cc.cellToString(c3);
		}
		XSSFCell c4 = ws.getRow(7).getCell(11);
		String cl4;
		if (c4 == null) {
			cl4 = "";
		} else {
			cl4 = cc.cellToString(c4);
		}
		XSSFCell c5 = ws.getRow(9).getCell(5);
		String cl5;
		if (c5 == null) {
			cl5 = "";
		} else {
			cl5 = cc.cellToString(c5);
		}
		XSSFCell c6 = ws.getRow(9).getCell(11);
		String cl6;
		if (c6 == null) {
			cl6 = "";
		} else {
			cl6 = cc.cellToString(c6);
		}
		XSSFCell c7 = ws.getRow(11).getCell(5);
		String cl7;
		if (c7 == null) {
			cl7 = "";
		} else {
			cl7 = cc.cellToString(c7);
		}
		XSSFCell c8 = ws.getRow(11).getCell(11);
		String cl8;
		if (c8 == null) {
			cl8 = "";
		} else {
			cl8 = cc.cellToString(c8);
		}
		XSSFCell c9 = ws.getRow(13).getCell(5);
		String cl9;
		if (c9 == null) {
			cl9 = "";
		} else {
			cl9 = cc.cellToString(c9);
		}
		XSSFCell c10 = ws.getRow(13).getCell(11);
		String cl10;
		if (c10 == null) {
			cl10 = "";
		} else {
			cl10 = cc.cellToString(c10);
		}
		XSSFCell c11 = ws.getRow(15).getCell(5);
		String cl11;
		if (c11 == null) {
			cl11 = "";
		} else {
			cl11 = cc.cellToString(c11);
		}
		XSSFCell c12 = ws.getRow(17).getCell(5);
		String cl12;
		if (c12 == null) {
			cl12 = "";
		} else {
			cl12 = cc.cellToString(c12);
		}
		XSSFCell c13 = ws.getRow(17).getCell(11);
		String cl13;
		if (c13 == null) {
			cl13 = "";
		} else {
			cl13 = cc.cellToString(c13);
		}
		XSSFCell c14 = ws.getRow(19).getCell(5);
		String cl14;
		if (c14 == null) {
			cl14 = "";
		} else {
			cl14 = cc.cellToString(c14);
		}
		XSSFCell c15 = ws.getRow(21).getCell(5);
		String cl15;
		if (c15 == null) {
			cl15 = "";
		} else {
			cl15= cc.cellToString(c15);
		}
		XSSFCell c16 = ws.getRow(21).getCell(11);
		String cl16;
		if (c16 == null) {
			cl16 = "";
		} else {
			cl16 = cc.cellToString(c16);
		}
		XSSFCell c17 = ws.getRow(23).getCell(5);
		String cl17;
		if (c17 == null) {
			cl17 = "";
		} else {
			cl17 = cc.cellToString(c17);
		}
		XSSFCell c18 = ws.getRow(25).getCell(5);
		String cl18;
		if (c18 == null) {
			cl18 = "";
		} else {
			cl18 = cc.cellToString(c18);
		}
		XSSFCell c19 = ws.getRow(27).getCell(5);
		String cl19;
		if (c19 == null) {
			cl19 = "";
		} else {
			cl19 = cc.cellToString(c19);
		}
		XSSFCell c20 = ws.getRow(27).getCell(11);
		String cl20;
		if (c20 == null) {
			cl20 = "";
		} else {
			cl20 = cc.cellToString(c20);
		}
		XSSFCell c21 = ws.getRow(29).getCell(5);
		String cl21;
		if (c21 == null) {
			cl21 = "";
		} else {
			cl21 = cc.cellToString(c21);
		}
		XSSFCell c22 = ws.getRow(31).getCell(5);
		String cl22;
		if (c22 == null) {
			cl22 = "";
		} else {
			cl22 = cc.cellToString(c22);
		}
		XSSFCell c23 = ws.getRow(33).getCell(5);
		String cl23;
		if (c23 == null) {
			cl23 = "";
		} else {
			cl23 = cc.cellToString(c23);
		}
		fis.close();
		
		/*
		 * Created By: NhanVH
		 * Description: Handle Checklist 1
		 */
		
		ArrayList<String> arr = new ArrayList<String>();
		if (cl1.length() == 0 || cl2.length() == 0 || cl3.length() == 0 || cl4.length() == 0 || cl5.length() == 0
				|| cl6.length() == 0 || cl7.length() == 0 || cl8.length() == 0 || cl9.length() == 0
				|| cl10.length() == 0 || cl11.length() == 0 || cl12.length() == 0 || cl13.length() == 0
				|| cl14.length() == 0 || cl15.length() == 0 || cl16.length() == 0 || cl17.length() == 0
				|| cl18.length() == 0 || cl19.length() == 0 || cl19.length() == 0 || cl20.length() == 0
				|| cl21.length() == 0 || cl22.length() == 0 || cl23.length() != 0)
			arr.add(ws.getSheetName());
		if (cl1.length() == 0)
//			ws.getRow(5).createCell(5).setCellStyle(cellStyle);
			arr.add("/F6");
		if (cl2.length() == 0)
//			ws.getRow(5).createCell(11).setCellStyle(cellStyle);
			arr.add("/L6");
		if (cl3.length() == 0)
//			ws.getRow(7).createCell(5).setCellStyle(cellStyle);
			arr.add("/F8");
		if (cl4.length() == 0)
//			ws.getRow(7).createCell(11).setCellStyle(cellStyle);
			arr.add("/L8");
		if (cl5.length() == 0)
//			ws.getRow(9).createCell(5).setCellStyle(cellStyle);
			arr.add("/F10");
		if (cl6.length() == 0)
//			ws.getRow(9).createCell(11).setCellStyle(cellStyle);
			arr.add("/L10");
		if (cl7.length() == 0)
//			ws.getRow(11).createCell(5).setCellStyle(cellStyle);
			arr.add("/F12");
		if (cl8.length() == 0)
//			ws.getRow(11).createCell(11).setCellStyle(cellStyle);
			arr.add("/L12");
		if (cl9.length() == 0)
//			ws.getRow(13).createCell(5).setCellStyle(cellStyle);
			arr.add("/F14");
		if (cl10.length() == 0)
//			ws.getRow(13).createCell(11).setCellStyle(cellStyle);
			arr.add("/L14");
		if (cl11.length() == 0)
//			ws.getRow(15).createCell(5).setCellStyle(cellStyle);
			arr.add("/F16");
		if (cl12.length() == 0)
//			ws.getRow(17).createCell(5).setCellStyle(cellStyle);
			arr.add("/F18");
		if (cl13.length() == 0)
//			ws.getRow(17).createCell(11).setCellStyle(cellStyle);
			arr.add("/L18");
		if (cl14.length() == 0)
//			ws.getRow(19).createCell(5).setCellStyle(cellStyle);
			arr.add("/F20");
		if (cl15.length() == 0)
//			ws.getRow(21).createCell(5).setCellStyle(cellStyle);
			arr.add("/F22");
		if (cl16.length() == 0)
//			ws.getRow(21).createCell(11).setCellStyle(cellStyle);
			arr.add("/L22");
		if (cl17.length() == 0)
//			ws.getRow(23).createCell(5).setCellStyle(cellStyle);
			arr.add("/F24");
		if (cl18.length() == 0)
//			ws.getRow(25).createCell(5).setCellStyle(cellStyle);
			arr.add("/F26");
		if (cl19.length() == 0)
//			ws.getRow(27).createCell(5).setCellStyle(cellStyle);
			arr.add("/F28");
		if (cl20.length() == 0)
//			ws.getRow(27).createCell(11).setCellStyle(cellStyle);
			arr.add("/L28");
		if (cl21.length() == 0)
//			ws.getRow(29).createCell(5).setCellStyle(cellStyle);
			arr.add("/F30");
		if (cl22.length() == 0)
//			ws.getRow(31).createCell(5).setCellStyle(cellStyle);
			arr.add("/F32");
		if (cl23.length() != 0)
//			ws.getRow(33).createCell(5).setCellStyle(cellStyle);
			arr.add("/F34");
		String result = "";
		for(int i = 0; i < arr.size(); i++) {
			result = result + arr.get(i);
		}
		
		if(arr.size() > 0) {
			ws1.getRow(1).createCell(4).setCellValue(result);
			ws1.getRow(1).getCell(4).setCellStyle(cellStyle1);
		}else {
			arr.clear();
		}
		
		/*
		 * Created By: NhanVH
		 * Description: Handle checklist 2
		 */
		if (cl12.equals("不要")) {
			if (cl13 != String.valueOf(0)) {
				ws1.getRow(2).createCell(4).setCellValue("L18");
				ws1.getRow(2).getCell(4).setCellStyle(cellStyle1);
			}
		} else {
			ws1.getRow(2).createCell(4).setCellValue("F18 & L18");
			ws1.getRow(2).getCell(4).setCellStyle(cellStyle1);
		}
		
		/*
		 * Created By: NhanVH
		 * Description: Handle checklist 3,4,5,6
		 */
		XSSFSheet ws2 = wb.getSheetAt(2);
		XSSFCell c24 = ws2.getRow(1).getCell(4);
		String cl24;
		if (c24 == null) {
			cl24 = "";
		} else {
			cl24 = cc.cellToString(c24);
		}
		XSSFCell c25 = ws2.getRow(2).getCell(4);
		String cl25;
		if (c25 == null) {
			cl25 = "";
		} else {
			cl25 = cc.cellToString(c25);
		}
		
		XSSFCell c26 = ws2.getRow(1).getCell(5);
		String cl26;
		if (c26 == null) {
			cl26 = "";
		} else {
			cl26 = cc.cellToString(c26);
		}
		XSSFCell c27 = ws2.getRow(2).getCell(5);
		String cl27;
		if (c27 == null) {
			cl27 = "";
		} else {
			cl27 = cc.cellToString(c27);
		}
		ArrayList<String> arr2 = new ArrayList<String>();
		if(!cl9.equals("2画面（事前・事後アンケート有）")) {
			if(ws2.getRow(1).getCell(3).toString().equals("✔")) {
				if(!cl24.equals("1.0")) {
					ws1.getRow(3).createCell(4).setCellValue(ws2.getSheetName() + "/E2");
					ws1.getRow(3).getCell(4).setCellStyle(cellStyle1);
				}else {
					if(!cl26.equals("事前")) {
						ws1.getRow(5).createCell(4).setCellValue(ws2.getSheetName() + "/F2");
						ws1.getRow(5).getCell(4).setCellStyle(cellStyle1);
					}
				}
			}else {
				ws1.getRow(3).createCell(4).setCellValue(ws2.getSheetName() + "/D2 & E2");
				ws1.getRow(3).getCell(4).setCellStyle(cellStyle1);
			}
		}else {
			if(ws2.getRow(1).getCell(3).toString().equals("✔")) {
				if(!cl24.equals("1.0")) {
					arr2.add("/E2");
				}else {
					if(!cl26.equals("事前")) {
						ws1.getRow(5).createCell(4).setCellValue(ws2.getSheetName() + "/F2");
						ws1.getRow(5).getCell(4).setCellStyle(cellStyle1);
					}
				}
			}else {
					arr2.add("/D2 & E2");
			}
			if(ws2.getRow(2).getCell(3).toString().equals("✔")) {
				if(!cl25.equals("2.0")) {
					arr2.add("/E3");
				}else {
					if(!cl27.equals("事後")) {
						ws1.getRow(6).createCell(4).setCellValue(ws2.getSheetName() + "/F3");
						ws1.getRow(6).getCell(4).setCellStyle(cellStyle1);
					}
				}
			}else {
					arr2.add("/D3 & E3");
			}
			if(arr2.size() > 0) {
				arr2.add(0, ws2.getSheetName());
			}
			
			String result2 = "";
			for(int i = 0; i < arr2.size(); i++) {
				result2 = result2 + arr2.get(i);
			}
			
			if(arr2.size() > 0) {
				ws1.getRow(4).createCell(4).setCellValue(result2);
				ws1.getRow(4).getCell(4).setCellStyle(cellStyle1);
			}
		}
		
		/*
		 * Created By: NhanVH
		 * Description: Handle checklist 7,8,9
		 */
		XSSFCell c28 = ws2.getRow(1).getCell(6);
		String cl28;
		if (c28 == null) {
			cl28 = "";
		} else {
			cl28 = cc.cellToString(c28);
		}
		
		XSSFCell c29 = ws2.getRow(1).getCell(7);
		String cl29;
		if (c29 == null) {
			cl29 = "";
		} else {
			cl29 = cc.cellToString(c29);
		}
		
		XSSFCell c30 = ws2.getRow(1).getCell(8);
		String cl30;
		if (c30 == null) {
			cl30 = "";
		} else {
			cl30 = cc.cellToString(c30);
		}
		
		XSSFCell c31 = ws2.getRow(1).getCell(9);
		String cl31;
		if (c31 == null) {
			cl31 = "";
		} else {
			cl31 = cc.cellToString(c31);
		}
		
		XSSFCell c32 = ws2.getRow(2).getCell(6);
		String cl32;
		if (c32 == null) {
			cl32 = "";
		} else {
			cl32 = cc.cellToString(c32);
		}
		
		if(cl24.equals("1.0")) {
			ArrayList<String> arr3 = new ArrayList<String>();
			if(cl26.length() == 0 || cl28.length() == 0 || cl29.length() == 0 || cl30.length() == 0 || cl31.length() == 0) {
				arr3.add(ws2.getSheetName());
			}
			if(cl26.length() == 0) {
				arr3.add("/F2");
			}
			if(cl28.length() == 0) {
				arr3.add("/G2");
			}
			if(cl29.length() == 0) {
				arr3.add("/H2");
			}
			if(cl30.length() == 0) {
				arr3.add("/I2");
			}
			if(cl31.length() == 0) {
				arr3.add("/J2");
			}
			
			String result3 = "";
			for(int i = 0; i < arr3.size(); i++) {
				result3 = result3 + arr3.get(i);
			}
			
			if(arr3.size() > 0) {
				ws1.getRow(7).createCell(4).setCellValue(result3);
				ws1.getRow(7).getCell(4).setCellStyle(cellStyle1);
			}
		}
		
		if(cl25.equals("2.0")) {
			if(cl32.length() == 0) {
				ws1.getRow(8).createCell(4).setCellValue(ws2.getSheetName() + "/G3");
				ws1.getRow(8).getCell(4).setCellStyle(cellStyle1);
			}
		}
		String str = "PJコード_TOP画像.jpg";
		if(!cl31.toLowerCase().contains(str.toLowerCase())) {
			ws1.getRow(9).createCell(4).setCellValue(ws2.getSheetName() + "/J2");
			ws1.getRow(9).getCell(4).setCellStyle(cellStyle1);
		}
		/*
		 * Created By: NhanVH
		 * Description: Execute record data on file and close file.
		 */
		
		FileOutputStream fops = new FileOutputStream(new File(input));
		wb.write(fops);
		fops.close();
		wb.close();
		FileOutputStream fops1 = new FileOutputStream(new File(output));
		wb1.write(fops1);
		fops1.close();
		wb1.close();
		
		System.out.println("The program Ended, Please check again!");
	}
}
