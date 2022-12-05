package projects;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class DemoProfileLoginCheck {

	static WebDriver driver;

	public static void setup() {

		System.setProperty("webdriver.chrome.driver",
				"E://CS//Selenium//Chrome Driver//chromedriver_win32/chromedriver.exe");
		driver = new ChromeDriver();                   
		driver.manage().window().maximize();			
		driver.get("https://itera-qa.azurewebsites.net/Login"); 
	}

	public static void main(String[] args) throws IOException, InterruptedException {

		setup();

		String expectedTitle = "Testautomation practice page";
		String acctualTitle = driver.getTitle();
		System.out.println("Titel is" + acctualTitle);
		if (acctualTitle == expectedTitle) {
			System.out.print("Title Matched!!");
		} else {
			System.out.println("Failed!!");
		}

		File file = new File("E:\\CS\\Java\\Java Project\\SeleniumProjects\\Excel\\data.xlsx");
		FileInputStream inputstream = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(inputstream);
		XSSFSheet sheet = wb.getSheet("Sheet1");
		// XSSFRow row=sheet.getRow(0);
		XSSFRow row = null;
		XSSFCell cell = null;
		String userName = null;
		String password = null;
		// System.out.println(sheet.getLastRowNum());
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			for (int j = 0; j <= row.getLastCellNum(); j++) {
				cell = row.getCell(j);
				if (j == 0) {
					userName = cell.getStringCellValue();
				}
				if (j == 1) {
					password = cell.getStringCellValue();
				}

			}

			driver.findElement(By.name("Username")).sendKeys(userName);
			driver.findElement(By.name("Password")).sendKeys(password);
			driver.findElement(By.name("login")).click();
			// driver.findElement(By.xpath("//a[normalize-space()='Log out']")).click();

			String result = null;
			try {
				Boolean isLoggedIn = driver.findElement(By.xpath("//a[normalize-space()='Log out']")).isDisplayed();
				if (isLoggedIn == true) {
					result = "PASS";
				}
				System.out.println("User Name : " + userName + " ----  " + "Password : " + password
						+ "----- Login success ? ------ " + result);

				driver.findElement(By.xpath("//a[normalize-space()='Log out']")).click();
			} catch (Exception e) {
				Boolean isError = driver
						.findElement(By.xpath("//label[normalize-space()='Wrong username or password']")).isDisplayed();
				if (isError == true) {
					result = "FAIL";
				}
				System.out.println("User Name : " + userName + " ----  " + "Password : " + password
						+ "----- Login success ? ------ " + result);
			}
			Thread.sleep(1000);
			driver.findElement(By.name("login")).click();
		}
		Thread.sleep(3000);
		driver.close();
	}

}
