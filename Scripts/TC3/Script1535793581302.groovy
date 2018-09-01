import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook

import internal.GlobalVariable as GlobalVariable

/**
 * This Test Case will output a Excel file which contains a sheet named `TC3` which
 * contains a pangram.
 *
 * Apache POI HSSFWorkbook objct will be prepared by MyTestListener@BeforeTestSuite method
 * or MyTestListener@BeforeTestCase method.
 */

HSSFWorkbook wb = (HSSFWorkbook)GlobalVariable.WORKBOOK
String sheetName = (String)GlobalVariable.CURRENT_TESTCASE_NAME
HSSFSheet sheet = CustomKeywords.'com.kazurayam.ksbackyard.HSSFSupport.findSheet'(wb, sheetName)

HSSFRow row = sheet.createRow(0)
row.createCell(0).setCellValue('El veloz murciélago hindú comía feliz cardillo y kiwi. La cigüeña tocaba el saxofón detrás del palenque de paja')
