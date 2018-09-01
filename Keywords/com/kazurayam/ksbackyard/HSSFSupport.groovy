package com.kazurayam.ksbackyard

import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook


class HSSFSupport {

	/**
	 *
	 * @param workbook
	 * @param sheetName
	 * @return
	 */
	static HSSFSheet findSheet(HSSFWorkbook workbook, String sheetName) {
		if (workbook.getSheet(sheetName) != null) {
			return workbook.getSheet(sheetName)
		} else {
			return workbook.createSheet(sheetName)
		}
	}
}