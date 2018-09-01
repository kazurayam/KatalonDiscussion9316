package com.kazurayam.ksbackyard

import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook

/**
 * 
 * 
 * @author kazurayam
 *
 */
class HSSFSupport {

	/**
	 * This static method looks up a sheet with the specified name in the workbook.
	 * If a sheet is found, then it will be returned.
	 * If no sheet with the name is found, the a new sheet will be created and returned.
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