package Package;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Class {

	private static Workbook wb;
	private static Sheet sh;
	private static FileInputStream fis;
	private static FileOutputStream fos;
	private static Row row;
	private static Cell cell;

	static WebDriver driver;
	static String niids = "Names";
	static String new_niids = "";
	static String valueMissing = "0";
	static int a = 0, b = 0, c = 0;
	static int initial_scroll = 420;
	static int scroll_amount = 20;

	public static void main(String[] args) throws Exception {

		// Create Properties & FileInputStream references to work with configuration file
		Properties prop = new Properties();
		FileInputStream ipfs = new FileInputStream(
				"C:\\Users\\nshanabo\\Documents\\Eclipse Workspaces\\WS005\\ArtifactID\\src\\main\\java\\Package\\config.properties");
		prop.load(ipfs);
		
		niids = prop.getProperty("niids_number");
		
		// File name generation
		String path = "C:\\Users\\nshanabo\\Desktop\\Newfolder\\";
		String fNameHalf = "C-Pro_NIIDS_";
		String fNameFull = fNameHalf + "Names" + ".xlsx";
		String fileFullPath = path + fNameFull;

		// Excel file
		fis = new FileInputStream(fileFullPath);
		wb = WorkbookFactory.create(fis);
		sh = wb.getSheet("Sheet1");

		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\nshanabo\\Downloads\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);

		driver.get(prop.getProperty("url"));
		System.out.println("URL Loaded.");

		driver.findElement(By.id("user_username")).sendKeys(prop.getProperty("userName"));
		driver.findElement(By.id("user_password")).sendKeys(prop.getProperty("password"));
		driver.findElement(By.xpath("//button[contains(.,'Login')]")).submit();

		// Goto NIIDS search
		driver.findElement(By.xpath("//div[@id='bs-example-navbar-collapse-1']/ul/li/a/img")).click();
		driver.findElement(By.xpath("//input[@id='searchString']")).clear();
		driver.findElement(By.xpath("//input[@id='searchString']")).sendKeys(niids);
		driver.findElement(By.xpath("//div[@id='searchResult']/div/table/tbody/tr/td[2]/a/b")).click();
		driver.findElement(By.xpath("//input[@id='button']")).click();
		driver.findElement(By.xpath("//input[@id='filter_contract_statusDRAFT']")).click();
		
		int totalRecords = sh.getLastRowNum();
		
		for (int j = 1; j<= totalRecords; j++) {
			
			Row current_row = sh.getRow(j);
			new_niids=current_row.getCell(0).getStringCellValue();
			
			driver.findElement(By.xpath("//input[@id='search']")).clear();
			driver.findElement(By.xpath("//input[@id='search']")).sendKeys(new_niids);
			driver.findElement(By.xpath("//input[@id='button']")).click();

			
			// Get web-table row count
			List<WebElement> rowList = driver.findElements(By.xpath("//table[@id='requestTable']/tbody/tr"));
			int rowCount = rowList.size();

			// Loop through and get the data into excel file
			next_rec:
			for (int i = 1; i <= rowCount; i++) {
				valueMissing = "0";
				JavascriptExecutor js = (JavascriptExecutor) driver;
				js.executeScript("window.scrollBy(0,420)", "");

				// Get XPATH of Name field
				String firstHalfXpath_Name = "//table[@id='requestTable']/tbody/tr[";
				String secondHalfXpath_Name = "]/td[5]";
				String fullXpath_Name = firstHalfXpath_Name + i + secondHalfXpath_Name;

				if (driver.findElement(By.xpath(fullXpath_Name)).getText().equals(new_niids)) {
					System.out.println("Matched");
				}
				else {
					continue next_rec;
				}
				
				// Get XPATH of 'View' Button
				String firstHalfXpath_View = "//table[@id='requestTable']/tbody/tr[";
				String secondHalfXpath_View = "]/td[13]";
				String fullXpath_View = firstHalfXpath_View + i + secondHalfXpath_View;

				// Get element to be visible
				String s1 = firstHalfXpath_View + (i+1) + secondHalfXpath_View;

				// Setup the spreadsheet
				row = sh.createRow(j);
				cell = row.createCell(0);
				cell.setCellValue(new_niids);
				System.out.print(driver.findElement(By.xpath(fullXpath_Name)).getText() + "\t");
				
				 
				try {
					driver.findElement(By.xpath(fullXpath_View)).click();
					System.out.println("Scrool: Inside first try block.");
				}
				
				catch (Exception e2) {
					System.out.println("Scrool: Inside first catch block.");
					JavascriptExecutor jst = (JavascriptExecutor) driver;
					jst.executeScript("arguments[0].scrollIntoView(true);", driver.findElement(By.xpath(s1)));
					Thread.sleep(1000);
					jst.executeScript("window.scrollBy(0,-200)");
					Thread.sleep(1000);
					driver.findElement(By.xpath(fullXpath_View)).click();
				}
				
				// Get 'Signed-date' from 'Details' tab.
				driver.findElement(By.xpath("//a[contains(text(),'Details')]")).click(); // Details on left side menu
				JavascriptExecutor js1 = (JavascriptExecutor) driver;
				js1.executeScript("arguments[0].scrollIntoView(true);", driver.findElement(By.xpath("//div[@id='FLD_contract_signdate']")));
				cell = row.createCell(1);
				try {
					cell.setCellValue(driver.findElement(By.xpath("//div[@id='FLD_contract_signdate']")).getText());
					System.out.print(driver.findElement(By.xpath("//div[@id='FLD_contract_signdate']")).getText() + "\t");
					a = 1;
				}
				
				catch (Exception e3) {
					valueMissing = "1";
					System.out.println("Signed date could not be captured successfully.");
					cell.setCellValue("Data Not Available");
				}

				// Get 'Offer-date' From 'Requests' page
				JavascriptExecutor jst = (JavascriptExecutor) driver;
				jst.executeScript("window.scroll(0,0)");
				//window.scrollTo(0, 0);
				driver.findElement(By.partialLinkText("Request")).click();
				try {
					driver.findElement(By.xpath("//div[@id='FLD_contract_offerid']/span")).click();
					JavascriptExecutor js2 = (JavascriptExecutor) driver;
					js2.executeScript("arguments[0].scrollIntoView(true);", driver.findElement(By.xpath("//div[@id='FLD_offer_offerdate']")));
					cell = row.createCell(2);
					cell.setCellValue(driver.findElement(By.xpath("//div[@id='FLD_offer_offerdate']")).getText());
					System.out.print(driver.findElement(By.xpath("//div[@id='FLD_offer_offerdate']")).getText() + "\t");
					b = 1;
				}
				catch (Exception e1) {
					valueMissing = "1";
					cell = row.createCell(2);
					cell.setCellValue("Data Not Available");
					System.out.print("Data Not Available" + "\t");
				}
				

				// Get 'Publish Date' - from 'Next Screen'
				try {
					driver.findElement(By.xpath("//a[contains(text(),'Details')]")).click();
					driver.findElement(By.xpath("//a[contains(text(),'Show request')]")).click();
					cell = row.createCell(3);
					cell.setCellValue(driver.findElement(By.xpath("//div[@id='FLD_requisition_publisheddate']")).getText());
					System.out.print(driver.findElement(By.xpath("//div[@id='FLD_requisition_publisheddate']")).getText() + "\t");
					System.out.println("");
					c = 1;
				}
				catch (Exception ee) {
					cell = row.createCell(3);
					cell.setCellValue("Data Not Available");
					System.out.print("Data Not Available" + "\n");
				}
					
				
				// Click on 'Back' Button - and go to correct fetch screen for the next record.
				// Goto NIIDS search
				System.out.println("Place 001");
				try {
					driver.findElement(By.xpath("//div[@id='bs-example-navbar-collapse-1']/ul/li/a/img")).click();
					System.out.println("Place 002");
				}
				catch (Exception e4) {
					if (valueMissing == "1") {
						JavascriptExecutor jsa = (JavascriptExecutor)driver;
						jsa.executeScript("arguments[0].scrollIntoView();", driver.findElement(By.xpath("//div[@id='topMessageModal']/div/div/div/button/img"))); 
						driver.findElement(By.xpath("//div[@id='topMessageModal']/div/div/div/button/img")).click();
					}
					WebElement element = driver.findElement(By.xpath("//div[@id='bs-example-navbar-collapse-1']/ul/li/a/img"));
					JavascriptExecutor jsx = (JavascriptExecutor)driver;
					jsx.executeScript("arguments[0].scrollIntoView();", element); 
					Actions a2 = new Actions(driver);
					a2.moveToElement(driver.findElement(By.xpath("//div[@id='bs-example-navbar-collapse-1']/ul/li/a/img"))).click().build().perform();
					System.out.println("Place 003");
				}
				
				driver.findElement(By.xpath("//input[@id='searchString']")).clear();
				driver.findElement(By.xpath("//input[@id='searchString']")).sendKeys(niids);
				try {
					driver.findElement(By.xpath("//div[@id='searchResult']/div/table/tbody/tr/td[2]/a/b")).click();
					System.out.println("Place 004");
				}
				catch (Exception e5) {
					Actions a3 = new Actions(driver);
					a3.moveToElement(driver.findElement(By.xpath("//div[@id='searchResult']/div/table/tbody/tr/td[2]/a/b"))).click().build().perform();
					System.out.println("Place 005");
				}
				
				try {
					driver.findElement(By.xpath("//input[@id='button']")).click();
					System.out.println("Place 006");
				}
				catch (Exception e6) {
					Actions a4 = new Actions(driver);
					a4.moveToElement(driver.findElement(By.xpath("//input[@id='button']"))).click().build().perform();
					System.out.println("Place 007");
				}
				
				try {
					driver.findElement(By.xpath("//input[@id='filter_contract_statusDRAFT']")).click();
					System.out.println("Place 008");
				}
				catch (Exception e7) {
					Actions a5 = new Actions(driver);
					a5.moveToElement(driver.findElement(By.xpath("//input[@id='filter_contract_statusDRAFT']"))).click().build().perform();
					System.out.println("Place 009");
				}
				driver.findElement(By.xpath("//input[@id='search']")).clear();
				System.out.println("Old value should have been cleared.");
				driver.findElement(By.xpath("//input[@id='search']")).sendKeys(new_niids);
				
				try {
					driver.findElement(By.xpath("//input[@id='button']")).click();
					System.out.println("Place 010");
				}
				catch (Exception e8) {
					Actions a6 = new Actions(driver);
					a6.moveToElement(driver.findElement(By.xpath("//input[@id='button']"))).click().build().perform();
					System.out.println("Place 011");
				}
				
				System.out.println("valueMissing: " + valueMissing);
				
				fos = new FileOutputStream(fileFullPath);
				wb.write(fos);
				fos.flush();
				System.out.println("Done with this record. Fetching the next record.");
					
			} // End of loop - for

		} // End of outer for loop (variable j)
		
		System.out.println("Records written to the spreadsheet.");
		fos.close();
		System.out.println("Process completed.");

	} // End of method - main

} // End of class - Class