package internal

import com.kms.katalon.core.configuration.RunConfiguration
import com.kms.katalon.core.testobject.ObjectRepository as ObjectRepository
import com.kms.katalon.core.testdata.TestDataFactory as TestDataFactory
import com.kms.katalon.core.testcase.TestCaseFactory as TestCaseFactory
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase

/**
 * This class is generated automatically by Katalon Studio and should not be modified or deleted.
 */
public class GlobalVariable {
     
    /**
     * <p>Profile default : instance of HSSFWorkbook</p>
     */
    public static Object WORKBOOK
     
    /**
     * <p></p>
     */
    public static Object CURRENT_TESTCASE_NAME
     
    /**
     * <p></p>
     */
    public static Object CURRENT_TESTSUITE_NAME
     
    /**
     * <p></p>
     */
    public static Object REPORT_FOLDER_NAME
     

    static {
        def allVariables = [:]        
        allVariables.put('default', ['WORKBOOK' : null, 'CURRENT_TESTCASE_NAME' : '', 'CURRENT_TESTSUITE_NAME' : '', 'REPORT_FOLDER_NAME' : ''])
        
        String profileName = RunConfiguration.getExecutionProfile()
        
        def selectedVariables = allVariables[profileName]
        WORKBOOK = selectedVariables['WORKBOOK']
        CURRENT_TESTCASE_NAME = selectedVariables['CURRENT_TESTCASE_NAME']
        CURRENT_TESTSUITE_NAME = selectedVariables['CURRENT_TESTSUITE_NAME']
        REPORT_FOLDER_NAME = selectedVariables['REPORT_FOLDER_NAME']
        
    }
}
