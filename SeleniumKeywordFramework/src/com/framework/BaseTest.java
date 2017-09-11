package com.framework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Parameters;

import com.aventstack.extentreports.ExtentReporter;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.MediaEntityBuilder;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

public class BaseTest {
	
	protected FileInputStream fileInput;
	protected int totalRows;
	protected int totalTestRows;
	protected Sheet s;
	protected Sheet tSheet;
	protected Workbook w ;
	protected WebDriver driver;
	protected ExtentHtmlReporter htmlReporter;
	protected ExtentReports reports;
	protected ExtentTest logger;
	boolean isPresent;
	protected FileOutputStream fileOuput;
	WebElement el;
	String imagePath="src/Report-Output/SnapShot/Err.png";	
	
	public BaseTest(){
		
		htmlReporter = new ExtentHtmlReporter("src/Report-Output/STMExtentReport.html");
		reports = new ExtentReports();
		reports.attachReporter(htmlReporter);
		
		reports.setSystemInfo("Host Name", "Software Testing Material");
		reports.setSystemInfo("Environment", "Automation Testing");
		reports.setSystemInfo("User Name", "giridhar");
		
	}
	
	@BeforeTest
	@Parameters({"Browser"})
	public void setUp(String browser){
		
		
		
		
		
		if(browser.equals("chrome")){
			System.setProperty("webdriver.chrome.driver","src/DriverServer/Chrome/chromedriver.exe");
		 
			ChromeOptions options = new ChromeOptions();
			options.addArguments("enable-automation");
			driver = new ChromeDriver(options);
		 driver.manage().window().maximize();
		 
		}
		
		if(browser.equals("firefox")){
			
			System.setProperty("webdriver.gecko.driver","src/DriverServer/Gecko/geckodriver.exe");
			
			DesiredCapabilities capabilities = DesiredCapabilities.firefox();
			FirefoxOptions options = new FirefoxOptions();
						
			options.addPreference("log", "{level: trace}");
					
			capabilities.setCapability("marionette", false);
			capabilities.setCapability("moz:firefoxOptions", options);
					
			 driver = new FirefoxDriver();
			 //driver.manage().window().maximize();
		
			}
		 driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		//driver.get(url);
		 
		
		
	}
	
	@AfterTest
	public void closeBrowser(){
		
		reports.flush();
		
	    driver.quit();
	}
	
	public void writeResultToExcel(String masterDataSheet, String status, int i) {
		try{
			
			FileInputStream file = new FileInputStream(masterDataSheet);

            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet("TestSuite");
            sheet.getRow(i).getCell(4,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(status);
		
		    FileOutputStream outFile =new FileOutputStream(Constants.Master_Data_Sheet);
		    workbook.write(outFile);
            outFile.close();
		
		
		
		}
		
		catch(Exception ex){
			
			System.out.println(ex.getMessage());
			
		}
		
		
	}
	
	
	public void readDataFromTestDatal(String masterDataSheet, String distributedApp) throws IOException {
		// TODO Auto-generated method stub
		fileInput = new FileInputStream(masterDataSheet);
		
		 w = new XSSFWorkbook(fileInput);
		
		 tSheet = w.getSheet(distributedApp);
		
		totalTestRows = tSheet.getLastRowNum();
		
	}
	
	public void readDataFromSuiteExcel(String suiteFile,String sheet) throws IOException{
		
		fileInput = new FileInputStream(suiteFile);
		
		 w = new XSSFWorkbook(fileInput);
		
		 s = w.getSheet(sheet);
		
		totalRows = s.getLastRowNum();
		
	}
	




public boolean executeAction(String keyword,String locator2, String locatervalue2,
		String testData2) throws InterruptedException, IOException {
	if(locator2!=null){
	if(locator2.equalsIgnoreCase("ID")){
		
		By ID = By.id(locatervalue2);
		isPresent= IsElementPresent(ID);
		System.out.println(isPresent);
		if(isPresent){
		 el = driver.findElement(ID); 
		}
		else 
			return false;
		
	}
	else if(locator2.equalsIgnoreCase("xpath")){
		
		By xPath = By.xpath(locatervalue2);
		isPresent= IsElementPresent(xPath);
		System.out.println(isPresent);
		if(isPresent){
		 el = driver.findElement(xPath); 
		}
		
		else 
			return false;
		
	}
	}
	
	
	switch(keyword){
	
	
	case "NavigateURL":
		
		driver.get(testData2);
		logger.log(Status.PASS,MarkupHelper.createLabel("Naigated to URL " + testData2  + " succcesfully" ,ExtentColor.GREEN));
		break;
		
	case "enterText":
		
			if(el.isEnabled()){
			
				try{
					el.sendKeys(testData2);
					logger.log(Status.PASS,MarkupHelper.createLabel("sent text " + testData2  + " succcesfully" ,ExtentColor.GREEN));
				}
				catch(StaleElementReferenceException ex){
					
					logger.log(Status.FAIL, MarkupHelper.createLabel("element gone stale " + ex.getMessage(), ExtentColor.RED));
					takeSnapShot(driver, imagePath);
					logger.log(Status.INFO, "Screenshot from :").addScreenCaptureFromPath(imagePath);
				}
				catch(Exception ex){
					
					logger.log(Status.FAIL, MarkupHelper.createLabel("element not found " + ex.getMessage(), ExtentColor.RED));
					takeSnapShot(driver, imagePath);
					logger.log(Status.INFO, "Screenshot from :").addScreenCaptureFromPath(imagePath);
				}
				
				
          	}
			
			else{
				
				logger.log(Status.FAIL, MarkupHelper.createLabel("element is not enabled stale ", ExtentColor.RED));
				takeSnapShot(driver, imagePath);
				logger.log(Status.INFO, "Screenshot from :").addScreenCaptureFromPath(imagePath);
				
			}	
			
			break;
			
	case "clickElement":
		
		if(el.isEnabled()){
		
			System.out.println("inside click");
			Thread.sleep(5000);
			try{
				
				//driver.findElement(By.xpath(".//*[@id='tsf']/div[2]/div[3]/center/input[1]")).click();
				//Actions action = new Actions(driver);
				//action.moveToElement(el).click();
				el.click();
				
				//((JavascriptExecutor)driver).executeScript("arguments[0].click;",el);
				logger.log(Status.PASS,MarkupHelper.createLabel("element clicked succcesfully" ,ExtentColor.GREEN));
			}
			catch(StaleElementReferenceException ex){
				
				logger.log(Status.FAIL, MarkupHelper.createLabel("element gone stale " + ex.getMessage(), ExtentColor.RED));
				takeSnapShot(driver, imagePath);
				logger.log(Status.INFO, "Screenshot from :").addScreenCaptureFromPath(imagePath);
				
			}
			catch(Exception ex){
				
				logger.log(Status.FAIL, MarkupHelper.createLabel("element not found " + ex.getMessage(), ExtentColor.RED));
				takeSnapShot(driver, imagePath);
				logger.log(Status.INFO, "Screenshot from :").addScreenCaptureFromPath(imagePath);
			}
			
      	}
		
		else{
			
			logger.log(Status.FAIL, MarkupHelper.createLabel("element is not enabled state ", ExtentColor.RED));
			takeSnapShot(driver, imagePath);
			logger.log(Status.INFO, "Screenshot from :").addScreenCaptureFromPath(imagePath);
			
		}
			
			break;
			
			default: break;
	
			
	}
	return true;
	
}




public void takeSnapShot(WebDriver driver2, String string) throws IOException {
	
	try{
	File SrcFile= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
	File DestFile=new File(string);
	FileUtils.copyFile(SrcFile, DestFile);
	}
	catch(Exception ex){
		
		System.out.println(ex.getMessage());
		
	}
	
}

private boolean IsElementPresent(By locater) throws IOException {
	// TODO Auto-generated method stub
	try{
		WebDriverWait wait = new WebDriverWait(driver, 10);	
	    wait.until(ExpectedConditions.visibilityOfElementLocated(locater));
	    
	    return true;
	}
	catch(NoSuchElementException ex){			
		
		logger.log(Status.FAIL, MarkupHelper.createLabel("No such elment found" + ex.getMessage(), ExtentColor.RED));
		takeSnapShot(driver, imagePath);
		logger.log(Status.INFO, "Screenshot from :").addScreenCaptureFromPath(imagePath);
		return false;
	}
	
   catch(Exception ex){			
		
		logger.log(Status.FAIL, MarkupHelper.createLabel("No such elment found ionside main exception" + ex.getMessage(), ExtentColor.RED));
		takeSnapShot(driver, imagePath);
		logger.log(Status.INFO, "Screenshot from :").addScreenCaptureFromPath(imagePath);
		return false;
	}
	
	
	
}
}
