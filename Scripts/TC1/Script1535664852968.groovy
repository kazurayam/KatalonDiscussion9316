import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory as CheckpointFactory
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as MobileBuiltInKeywords
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testcase.TestCaseFactory as TestCaseFactory
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testdata.TestDataFactory as TestDataFactory
import com.kms.katalon.core.testobject.ObjectRepository as ObjectRepository
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WSBuiltInKeywords
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUiBuiltInKeywords
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable

import com.kms.katalon.core.configuration.RunConfiguration
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFCell

String excelFileName = "results.xls"

Path projectDir = Paths.get(RunConfiguration.getProjectDir())
Path tmp = projectDir.resolve('tmp')
Path excel = tmp.resolve(excelFileName)
Files.createDirectories(tmp)
WebUI.comment("Excel file at ${excel.toString()}")

HSSFWorkbook wb
if (Files.exists(excel)) {
	FileInputStream fis = new FileInputStream(excel.toFile())
	wb = new HSSFWorkbook(fis)
} else {
	wb = new HSSFWorkbook()
}

String sheetName = "SHEET-0"
HSSFSheet sheet
if (wb.getSheet(sheetName) != null) {
	sheet = wb.getSheet(sheetName)
} else {
	sheet = wb.createSheet(sheetName)
}

HSSFRow row = sheet.createRow(0);
row.createCell(0).setCellValue(excelFileName);
row.createCell(1).setCellValue("foo");
row.createCell(2).setCellValue("bar");

FileOutputStream fileOut = new FileOutputStream(excel.toFile());
wb.write(fileOut);
fileOut.flush();
fileOut.close();