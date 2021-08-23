package pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.CacheLookup;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.PageFactory;
import org.testng.Assert;
import utility.BaseUtiltiy;

public class HomePage extends BaseUtiltiy{

	WebDriver Browser;
	public HomePage(WebDriver Browser) {
		this.Browser=Browser;
		PageName=this.getClass().getSimpleName();
	}
	@FindBy(how=How.XPATH, using="//div[@aria-label='Account']")
	@CacheLookup
	public WebElement AccountMenu;

	@FindBy(how=How.ID_OR_NAME, using="userNavigationLabel")
	@CacheLookup
	public WebElement LogoutMenu;

	@FindBy(how=How.XPATH, using="//span[contains(text(),'Log Out')]")
	@CacheLookup
	public WebElement Logout;

	public LoginPage Logout(){
		try {
			AccountMenu.click();
			Thread.sleep(2000);
			if (Logout.isDisplayed()) {
				Logout.click();
			}
		} catch (Exception e) {
			Testfail=true;
		}
		int Logout = Browser.findElements(By.name("login")).size();

		if (Logout==1) {
			return PageFactory.initElements(Browser, LoginPage.class);
		}
		else {
			Testfail=true;
			Assert.assertFalse(Testfail, "Test Case '"+TestCaseName+"', Page '"+PageName+"' for Data set '"+(DataSet+1)+"' got Fail");
			return null;
		}

	}
}
