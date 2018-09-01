import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook

import internal.GlobalVariable as GlobalVariable

/**
 * This Test Case will output a Excel file which contains a sheet named `TC3`.
 * The sheet contains a pangram in Spanish language.
 *
 * Apache POI HSSFWorkbook objct will be prepared by MyTestListener's @BeforeTestSuite method
 * or MyTestListener's @BeforeTestCase method.
 */

HSSFWorkbook wb = (HSSFWorkbook)GlobalVariable.WORKBOOK

// In the workbook, you can do anything you like

def updateSheet(HSSFWorkbook wb, String testCaseName, String sentence) {
	HSSFSheet sheet = CustomKeywords.'com.kazurayam.ksbackyard.HSSFSupport.findSheet'(wb, testCaseName)
	HSSFRow row = sheet.createRow(1)
	row.createCell(0).setCellValue(sentence)
}

// refer to the sheet 'TC1' created by the Test Case 'TC1' and update it
updateSheet(wb, 'TC1', 'The Test Case TC3 says Hello')

// refer to the sheet 'TC2' created by the Test Case 'TC2' and update it
updateSheet(wb, 'TC2', 'ТС3 говорил До свидания.')

/*
 * The Excel file will be written by MyTestListener's @AfterTestCase method.
 */
