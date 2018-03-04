package com.nhanvh.tools;

import org.apache.poi.xssf.usermodel.XSSFCell;

public class CheckCell {
	@SuppressWarnings("deprecation")
	public static String cellToString(XSSFCell cell) {
		int type;
		Object result;
		type = cell.getCellType();
		switch (type) {
		case 0:
			result = cell.getNumericCellValue();
			break;

		case 1:
			result = cell.getStringCellValue();
			break;

		default:
			result = "";
		}
		return result.toString();
	}
}
