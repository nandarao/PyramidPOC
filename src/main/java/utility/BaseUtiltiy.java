package utility;

import org.openqa.selenium.WebDriver;
import pages.HomePage;
import pages.LoginPage;

public class BaseUtiltiy {
	private static Object FileUtils;
	public WebDriver Browser;
	public  static 	ReadAndWrite_xlsx FilePath;
	public  static	final String TestCaseSheetName="Test Case List";
	public  static	final String TestDataSheetName="Test Data";
	public  static  final String Header ="Test Case List";
	public  static  String  TestCaseName = null;
	public  static	final String colName = "Runmode";
	public  static	String TestCaseRunmode=null;
	public  static	String TestRunModeTeatData[]= null;
	public  static	final String StatusUpdate = "Pass/Fail/Skip";
	public  static boolean TestCasePass=true;
	public  static boolean Testskip=false;
	public  static boolean Testfail=false;
	//public  static boolean isLogingIn=false;
	public  static int DataSet=-1;

	//Pages Name declare hear
	public static String PageName = null;

	public LoginPage login =null;
	public HomePage home = null;

	public static ReadAndWrite_xlsx TestCaseListExcel= new ReadAndWrite_xlsx("C:\\Users\\kishorenm.PYRAMIDINDIA\\IdeaProjects\\poc\\src\\main\\java\\excel\\Test Cases.xlsx");

	public static  String[] checkRunModeOfTestData(ReadAndWrite_xlsx xlsx, String SheetName, String HeaderOrTastCaseName, String colName){
		return xlsx.retrieveToRunFlagTestData(SheetName, HeaderOrTastCaseName, colName);
	}

	public static  String checkRunModerOfTestCase(ReadAndWrite_xlsx xlsx,String SheetName, String  Header , String TestCaseName , String colName){
		return xlsx.retrieveToRunFlagTestCase(SheetName, Header, TestCaseName, colName);
	}

	public static Object[][] getTestDataValus(ReadAndWrite_xlsx xlsx, String SheetName, String HeaderOrTastCaseName ){
		return xlsx.retrieveTestDataValues(SheetName, HeaderOrTastCaseName);

	}

	public static boolean writeExcutionStatusToTestData(ReadAndWrite_xlsx xlsx, String SheetName, String HeaderOrTastCaseName, String colName , int rowNumber , String Result){
		return	xlsx.writeExciquteResultInTestData(SheetName, HeaderOrTastCaseName, colName, rowNumber, Result);

	}

	public static boolean writeExcutionStatusToTestCase(ReadAndWrite_xlsx xlsx, String SheetName, String Header, String TastCaseName, String colName , String Result){
		return	xlsx.writeExciquteResultInTestCase(SheetName, Header, TastCaseName, colName, Result);

	}

//	public static void ScreenShots(WebDriver Browser, String PageName, String Result){
//
//		try {
//			File SS = ((TakesScreenshot) Browser)
//					.getScreenshotAs(OutputType.FILE);
//			FileUtils.copyFile(SS, new File("TestResult\\ScreenShots\\Error_in_Test Case "+ TestCaseName+"->"+PageName +"Round "+(DataSet+1)+ ".png"));
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//	}

	public static void setFileUtils(Object fileUtils) {
		FileUtils = fileUtils;
	}
}
