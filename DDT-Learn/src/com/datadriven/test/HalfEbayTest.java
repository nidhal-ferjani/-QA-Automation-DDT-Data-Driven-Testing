package com.datadriven.test;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.datadriven.data.DataDrivenProvider;

public class HalfEbayTest {

	private WebDriver webDriver;
	private String URL;

	@BeforeClass(enabled = true)
	@Parameters({ "env", "browser", "url" })
	public void setUp(@Optional("Test") String env, @Optional("Chrome") String browser, String URL) {

		if (browser.equalsIgnoreCase("chrome")) {
			System.setProperty("webdriver.chrome.driver", "D:\\NodeJS (WorkSpace)\\Home-Test-Automation"
					+ "\\Test-Automation-Framework-Egypte\\Drivers\\chromedriver.exe");
			webDriver = new ChromeDriver();

			webDriver.manage().window().maximize();
			webDriver.manage().deleteAllCookies();
			webDriver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
			webDriver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(10));
			this.URL = URL;

		} else {
			System.exit(0);
		}
	}

	@Test(dataProvider = DataDrivenProvider.INPUT_DATA_USER_EXCEL, dataProviderClass = DataDrivenProvider.class)
	public void testHalfEbayRegisterPage(String firstName, String lastName, String email, String password) {
		webDriver.get(URL);
		WebElement firstNameTxt = webDriver.findElement(By.id("firstname"));
		firstNameTxt.clear();
		firstNameTxt.sendKeys(firstName);
		
		WebElement lastNameTxt = webDriver.findElement(By.id("lastname"));
		lastNameTxt.clear();
		lastNameTxt.sendKeys(lastName);
		
		WebElement emailTxt = webDriver.findElement(By.id("Email"));
		emailTxt.clear();
		emailTxt.sendKeys(email);
		
		WebElement passwordTxt = webDriver.findElement(By.id("password"));
		passwordTxt.clear();
		passwordTxt.sendKeys(lastName);

	}

	@AfterClass(enabled = true)
	public void tearDown() { // Or cleanUp
		webDriver.quit();
	}

}
