package HomeProject;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.Select;

public class POM {
	WebDriver driver;
	Logger log = Logger.getLogger("POM");
	Select select;
	
	@FindBy(xpath = "//ul[@id='menu-all-departments-1']//child::a[@title='Laptops']")
	WebElement laptopButton;
	
	@FindBy(xpath = "//i[@class='departments-menu-v2-icon fa fa-list-ul']")
	WebElement productButton;
	
	@FindBy(xpath = "//ul[@id='menu-all-departments-1']//child::a[@title='Tablet/Watches']")
	WebElement tabletButton;
	
	@FindBy(xpath = "//div[@class='product-loop-body product-item__body']//child::h2[@class='woocommerce-loop-product__title']")
	List<WebElement> tabletNames;
	
	@FindBy(xpath = "//div[@class='product-loop-body product-item__body']//child::div[@class='product-short-description']")
	List<WebElement> tabletDescription;
	
	@FindBy(xpath = "//div[@class='product-loop-footer product-item__footer']//child::span[@class='woocommerce-Price-amount amount']")
	List<WebElement> tabletPrices;
	
	@FindBy(xpath = "//select[@name='ppp']")
	WebElement showAll;
	
	@FindBy(xpath = "//div[@class='product-thumbnail product-item__thumbnail']")
	List<WebElement> GetImages;
	
	public static void setCellData(String SheetName, int rowNum, int colNum, String inputData) throws Exception
	{
		String excelFilePath=".\\Documents\\Book1.xlsx";
		FileInputStream inputstream=new FileInputStream(excelFilePath);
		XSSFWorkbook excelBook = new XSSFWorkbook(inputstream);
		XSSFSheet sheet =  excelBook.getSheet(SheetName);	
		XSSFRow row = sheet.getRow(rowNum);
		if(row == null) {
			 row = sheet.createRow(rowNum);

		}
		XSSFCell cell = row.createCell(colNum);
		cell.setCellValue(inputData);
		FileOutputStream outputstream= new FileOutputStream(excelFilePath);
		excelBook.write(outputstream);
	}
	public void ShowAll() {
		select = new Select(showAll);
		select.selectByValue("-1");
		log.info("ShowAll is selected");

	}
	
	public void GetTabletNames() throws Exception {
		List<String> list1=new ArrayList<String>(); 
		for (WebElement i : tabletNames) {
			list1.add(i.getText());
			
		}
		for(int i = 0; i < list1.size();i++) {
			
			setCellData("Sheet1",i,0,list1.get(i));
		}
		log.info("Tablet Names Extracted");

	}
	public void GetTabletDescription() throws Exception {
		List<String> list1=new ArrayList<String>(); 
		for (WebElement i : tabletDescription) {
			list1.add(i.getText());
		}
		for(int i = 0; i < list1.size();i++) {
			setCellData("Sheet1",i,1,list1.get(i));
		}
		log.info("Tablet Description Extracted");

	}
	public void GetTabletPrices() throws Exception {
		List<String> list1=new ArrayList<String>(); 
		for (WebElement i : tabletPrices) {
			list1.add(i.getText());
		}
		for(int i = 0; i < list1.size();i++) {
			setCellData("Sheet1",i,2,list1.get(i));
		}
		log.info("Tablet Prices Extracted");

	}
	public void GetImages() throws AWTException, InterruptedException {
		for(WebElement pics:GetImages) {
			Actions act = new Actions(driver);
			act.sendKeys(Keys.PAGE_DOWN).build().perform();
			act.moveToElement(pics).contextClick().build().perform();
			Robot robot = new Robot();
			for (int i = 1; i <=8 ; i++) {
				robot.keyPress(KeyEvent.VK_DOWN);
				robot.keyRelease(KeyEvent.VK_DOWN);
			}
			robot.keyRelease(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
			Thread.sleep(3000);

			robot.keyRelease(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
			Thread.sleep(3000);



		}
	}
	
	
	public POM(WebDriver driver) {
		this.driver = driver;
		// This initElements method will create all WebElements
		PageFactory.initElements(driver, this);

	}
}
