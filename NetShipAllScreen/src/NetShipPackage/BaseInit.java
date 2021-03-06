package NetShipPackage;

import java.io.File;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Platform;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseInit {
	public static WebDriver Driver;

	@BeforeSuite
	public void beforeMethod() {
		DesiredCapabilities capabilities = new DesiredCapabilities();
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		// options.addArguments("headless");
		options.addArguments("headless");
		options.addArguments("--incognito");
		options.addArguments("--test-type");
		options.addArguments("--no-proxy-server");
		options.addArguments("--proxy-bypass-list=*");
		options.addArguments("--disable-extensions");
		options.addArguments("--no-sandbox");
		options.addArguments("--headless");
		options.addArguments("window-size=1366x788");
		capabilities.setPlatform(Platform.ANY);
		capabilities.setCapability(ChromeOptions.CAPABILITY, options);
		Driver = new ChromeDriver(options);
		// Default size
		Dimension currentDimension = Driver.manage().window().getSize();
		int height = currentDimension.getHeight();
		int width = currentDimension.getWidth();
		System.out.println("Current height: " + height);
		System.out.println("Current width: " + width);
		System.out.println("window size==" + Driver.manage().window().getSize());

		// Set new size
		Dimension newDimension = new Dimension(1366, 788);
		Driver.manage().window().setSize(newDimension);

		// Getting
		Dimension newSetDimension = Driver.manage().window().getSize();
		int newHeight = newSetDimension.getHeight();
		int newWidth = newSetDimension.getWidth();
		System.out.println("Current height: " + newHeight);
		System.out.println("Current width: " + newWidth);
	}

	@BeforeTest
	public void Login() throws Exception {
		WebDriverWait wait = new WebDriverWait(Driver, 50);
		System.out.println("**********Login Sucessfully**********");
		// ********************User Name and Password***********************
		// DEV
		// Driver.get("http://10.20.104.122:8075/login");
		// Staging
		Driver.get("http://stagingns.nglog.com/");
		wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.name("loginForm")));
		// Pre-Production
		// Driver.get("http://192.168.11.82:8074/");

		Driver.findElement(By.id("inputUsername")).clear();
		// DEV
		// Driver.findElement(By.id("inputUsername")).sendKeys("95008401");
		// Staging
		Driver.findElement(By.id("inputUsername")).sendKeys("automation");
		// Pre-Production
		// driver.findElement(By.id("inputUsername")).sendKeys("10327201");

		Driver.findElement(By.id("inputPassword")).clear();
		// DEV AND Staging
		Driver.findElement(By.id("inputPassword")).sendKeys("Auto@123");
		// Pre-Production
		// driver.findElement(By.id("inputPassword")).sendKeys("password");
		Driver.findElement(By.id("btnSignIn")).click();
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		System.out.println("**********Net Ship Information Popup**********");
		try {
			if (Driver.findElement(By.id("btnDismiss")).isDisplayed() == true) {
				getscreenshot("NetShipInfoPopup");
				Driver.findElement(By.id("btnDismiss")).click();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));
				System.out.println("Net Ship Info Pop up is display.");
			}
		} catch (Exception e) {
			System.out.println("Net Ship Info Pop up is not display.");
		}
		wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.id("ActiveOrderGrd")));
	}

	@AfterTest
	public void Logout() throws Exception {
		WebDriverWait wait = new WebDriverWait(Driver, 50);
		System.out.println("**********Logout**********");
		Driver.findElement(By.id("divUsername")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("hrefLogout")));
		Driver.findElement(By.id("hrefLogout")).click();
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));
	}

	@AfterSuite
	public void afterMethod() throws Exception {
		Driver.close();
		Driver.quit();
	}

	public void getscreenshot(String ScrSht) throws Exception {
		File scrFile = ((TakesScreenshot) Driver).getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(scrFile, new File("./Screenshots/" + ScrSht + ".jpg"));
	}

	public String CuDate() {
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy ");
		Date date = new Date();
		String date1 = dateFormat.format(date);
		System.out.println("Current Date :- " + date1);
		return date1;
	}

	public String CuDateTime() {
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH.mm");
		Date date = new Date();
		String date1 = dateFormat.format(date);
		System.out.println("Current Date :- " + date1);
		return date1;
	}

	public static String getDate(Calendar cal) {
		return "" + cal.get(Calendar.MONTH) + "/" + (cal.get(Calendar.DATE) + 1) + "/" + cal.get(Calendar.YEAR);
	}

	public void selectGroupBy() throws InterruptedException {
		WebDriverWait wait = new WebDriverWait(Driver, 50);
		Select SelectSort2 = new Select(Driver.findElement(By.id("drpGrouping")));
		SelectSort2.selectByIndex(1);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		String SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		String value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(2);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(3);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(4);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(5);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(6);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(7);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(8);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(9);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(0);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
	}
}
