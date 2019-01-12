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

// test case TC3 updates the sheet 'TC1'
updateSheet(wb, 'TC1', 'TC3 said Hello to TC1')

// test case TC3 updates the sheet 'TC2'
updateSheet(wb, 'TC2', 'ТС3 поприветствовал TC2')

/*
 * The Excel sheets will be saved into File by MyTestListener's @AfterTestCase method.
 */
