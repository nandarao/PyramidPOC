package TestCases;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.PageFactory;
import org.testng.Assert;
import org.testng.SkipException;
import org.testng.annotations.*;
import pages.LoginPage;
import utility.BaseUtiltiy;

import java.util.Hashtable;

public class LoginTest extends BaseUtiltiy{

	@BeforeTest
	public void checkTestCaseRunMode(){
		TestCaseName = this.getClass().getSimpleName();
		FilePath=BaseUtiltiy.TestCaseListExcel;
		TestCaseRunmode = BaseUtiltiy.checkRunModerOfTestCase(FilePath, TestCaseSheetName, Header, TestCaseName, colName);
//		System.out.println("value of runmode --------------"+TestCaseRunmode);
		if (!TestCaseRunmode.equalsIgnoreCase("y")) {
			BaseUtiltiy.writeExcutionStatusToTestCase(FilePath, TestCaseSheetName, TestCaseSheetName, TestCaseName, StatusUpdate, "Skip");
			throw  new SkipException("Test Case '"+TestCaseName+"' Rummode as 'N', So Skiping Excicution of '"+TestCaseName+ "'");
		}
	}


	@DataProvider
	public  Object[][] LoginTestData(){
		return BaseUtiltiy.getTestDataValus(FilePath, TestDataSheetName, TestCaseName);
	}

	@Test(dataProvider="LoginTestData")
	public void LoginTest(Hashtable<String, String>data){
		DataSet++;
		if(data.get("Runmode").equalsIgnoreCase("n")){
			Testskip=true;
			throw new SkipException("Test Case '"+TestCaseName+"', Test Data row number "+DataSet+" Runmode is No. So Skipping Its Execution.");
		}
		if (data.get("BrowserType").equalsIgnoreCase("CH")){
			ChromeOptions options = new ChromeOptions();
			options.addArguments("--disable-notifications");
			System.setProperty("webdriver.chrome.driver","src\\main\\java\\drivers\\chromedriver.exe");
			Browser = new ChromeDriver(options);

		}
		else if(data.get("BrowserType").equalsIgnoreCase("FF")){
			System.setProperty("webdriver.gecko.driver", "src\\main\\java\\drivers\\geckodriver.exe");
			DesiredCapabilities capabilities = DesiredCapabilities.firefox();
			capabilities.setCapability("marionette",true);
			Browser = new FirefoxDriver(capabilities);
		}
		Browser.manage().window().maximize();
		Browser.get(data.get("Url"));
		LoginPage lp=PageFactory.initElements(Browser, LoginPage.class);
		home= lp.doLogin(data.get("Username"), data.get("Password"));
		if(home!=null) Assert.assertTrue(home!=null, "Successfully Login");
		else Assert.assertTrue(home!=null, "Not able to Login");
		if(home.Logout()!=null) Assert.assertTrue(home!=null, "Successfully Logout");
		else Assert.assertTrue(home!=null, "Not able to Logout");
//		Assert.assertFalse(home!=null, "Test Case '"+TestCaseName+"', Page '"+PageName+"' for Data set '"+(DataSet+1)+"' got Fail -------> Not able to Logout");
	}

	@AfterMethod
	public void TestDataUpdateResult(){
		System.out.println("after method skip = "+Testskip + ", TestFail = "+Testfail+", TestCasePass ="+TestCasePass+" PageName "+PageName);

		if(Testskip){
			BaseUtiltiy.writeExcutionStatusToTestData(FilePath, TestDataSheetName, TestCaseName, StatusUpdate, DataSet+1, "Skip");
		}
		else if(Testfail){
			TestCasePass=false;
			BaseUtiltiy.writeExcutionStatusToTestData(FilePath, TestDataSheetName, TestCaseName, StatusUpdate, DataSet+1, "Fail");
			Browser.close();
		}
		else {
			BaseUtiltiy.writeExcutionStatusToTestData(FilePath, TestDataSheetName, TestCaseName, StatusUpdate, DataSet+1, "Pass");
			Browser.close();
		}
		Testskip=false;
		Testfail=false;
	}

	@AfterTest
	public void CloseBrowser(){
		if (TestCasePass) {
			BaseUtiltiy.writeExcutionStatusToTestCase(FilePath, TestCaseSheetName, Header, TestCaseName, StatusUpdate, "Pass");
		}
		else {
			BaseUtiltiy.writeExcutionStatusToTestCase(FilePath, TestCaseSheetName, Header, TestCaseName, StatusUpdate, "Fail");		}
	}
}
