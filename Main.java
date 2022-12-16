package HomeProject;

import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;



public class Main {
	WebDriver driver;
	POM obj;
	@BeforeTest
	public void Setup() {
		System.setProperty("webdriver.chrome.driver",
				"C:\\\\Users\\\\4397\\\\Documents\\\\SQA\\\\chromedriver_win32\\\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://galaxy.pk/");
		obj = new POM(driver);

		PropertyConfigurator.configure("log4j.properties");
		
	}
	@Test
	public void Test1() throws InterruptedException {
		obj.laptopButton.click();
		
	}
	@Test(dependsOnMethods = {"Test1"})
	public void Test2() throws InterruptedException {
		obj.ShowAll();

	}
		
	@Test(dependsOnMethods = {"Test1","Test2"})
	public void Test3() throws Exception {
		obj.GetTabletNames();

	}
	@Test(dependsOnMethods = {"Test1","Test2","Test3"})
	public void Test4() throws Exception {
		obj.GetTabletDescription();

	}
	@Test(dependsOnMethods = {"Test1","Test2","Test3","Test4"})
	public void Test5() throws Exception {
		obj.GetTabletPrices();

	}
	@Test(dependsOnMethods = {"Test1","Test2","Test3","Test4","Test5"})
	public void Test6() throws Exception {
		obj.GetImages();

	}
	
}
