import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook

import internal.GlobalVariable as GlobalVariable

/**
 * This Test Case will output a Excel file which contains a sheet named `TC1`.
 * The sheed contains a pangram in English language.
 * 
 * Apache POI HSSFWorkbook objct will be prepared by MyTestListener's @BeforeTestSuite method
 * or MyTestListener's @BeforeTestCase method.
 */

HSSFWorkbook wb = (HSSFWorkbook)GlobalVariable.WORKBOOK

// In the workbook, you can do anything you like 

// create a sheet named 'TC1'
String sheetName = (String)GlobalVariable.CURRENT_TESTCASE_NAME
HSSFSheet sheet = CustomKeywords.'com.kazurayam.ksbackyard.HSSFSupport.findSheet'(wb, sheetName)

// In the Cell `A1`, write a sentence
HSSFRow row = sheet.createRow(0)
row.createCell(0).setCellValue('The quick brown fox jumps over the lazy dog')

/*
 * The Excel file will be writte by MyTestListener's @AfterTestCase method.
 *
 */
