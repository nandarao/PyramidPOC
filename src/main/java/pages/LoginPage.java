package pages;

import java.io.File;

import org.openqa.selenium.*;
import org.openqa.selenium.support.CacheLookup;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.PageFactory;
import org.testng.Assert;
import org.testng.asserts.Assertion;

import utility.BaseUtiltiy;

public class LoginPage extends BaseUtiltiy {

WebDriver Browser;
	
	public LoginPage (WebDriver Browser){
		this.Browser=Browser;
		 PageName = this.getClass().getSimpleName();
	}
	
		
	@FindBy(how=How.XPATH, using=".//*[@id='u_0_p']")
	@CacheLookup
	public WebElement SelectLang;
	
	@FindBy(how = How.NAME, using="email")
	@CacheLookup
	public WebElement EmailOrPhone;
	
	@FindBy(how = How.NAME, using="pass")
	@CacheLookup
	public WebElement password;
	
	@FindBy(how = How.XPATH, using=".//*[@id='login_form']/div[1]/div[1]")
	@CacheLookup
	public WebElement ErrorLoing;	

	@FindBy(how = How.NAME, using="login")
	@CacheLookup
	public WebElement LogIn;
	
	public void emailOrPhone(String emailOrPhone){
	
		EmailOrPhone.sendKeys(emailOrPhone);
}

	public void password(String pwd){
		password.sendKeys(pwd);
}

	public void login(){
		LogIn.click();
}

 	public HomePage doLogin(String userIDValue, String passwordValue){
	try {
		System.out.println("Start TestSkip = "+Testskip + ", TestFail = "+Testfail+", TestCasePass ="+TestCasePass+ ", PageName "+PageName);
		emailOrPhone(userIDValue);
 		password(passwordValue);
 		if (LogIn.isDisplayed()) {
 			System.out.println("take screenshots");
 		}
 		login();

	} catch (Exception e) {
		Testfail=true;
		Assert.assertFalse(Testfail, "Test Case '"+TestCaseName+"', Page '"+PageName+"' for Data set '"+(DataSet+1)+"' got Fail");
	}
	int Logout = Browser.findElements(By.xpath("//div[@aria-label='Account']")).size();
	
	if (Logout==1) {
 		return PageFactory.initElements(Browser, HomePage.class);
	}
	else {
		Testfail=true;
		Assert.assertFalse(Testfail, "Test Case '"+TestCaseName+"', Page '"+PageName+"' for Data set '"+(DataSet+1)+"' got Fail");
		return null;
	}
	}
}
