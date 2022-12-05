package projects;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ProjectB {

	public static void main(String[] args) throws InterruptedException, IOException {

		System.setProperty("webdriver.chrome.driver",
				"E://CS//Selenium//Chrome Driver//chromedriver_win32/chromedriver.exe");
		WebDriver driver = new ChromeDriver();// open the browser
		driver.manage().window().maximize();
		driver.get("https://www.google.com");
		driver.findElement(By.xpath("//a[contains(text(),'English')]")).click();

		File file = new File("E:\\CS\\Java\\Java Project\\SeleniumProjects\\Excel\\Book1.xlsx");
		FileInputStream inputstream = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(inputstream);
		XSSFSheet sheet = wb.getSheet("Sheet1");

		XSSFRow row = null;
		XSSFCell cell = null;
		String keyword = null;
		// String[] array2= new String[sheet.getLastRowNum()];
		// System.out.println(sheet.getLastRowNum());
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			for (int j = 0; j <= row.getLastCellNum(); j++) {
				cell = row.getCell(j);
				if (j == 0) {
					keyword = cell.getStringCellValue();
				}
			}
			// System.out.print(keyword);

			driver.findElement(By.name("q")).sendKeys(keyword);
			Thread.sleep(2000);
			List<WebElement> listOfElements = driver.findElements(By.xpath("//ul/li[@role='presentation']"));
			int elementSize = listOfElements.size();
			String[] array = new String[elementSize];
			int a = 0;
			int maxLength = 0;
			String longestString = "";

			for (WebElement in : listOfElements) {
				array[a] = in.getText();
				a++;
				// System.out.println("Suggestion text:" + i.getText()+ "...." + "size of sen:"
				// + i.getText().length() );//name of the element from list
			}

			for (String s : array) {
				if (s.length() > maxLength) {
					maxLength = s.length();
					longestString = s;
				}
			}
			cell = row.createCell(1);
			cell.setCellValue(longestString);

			String smallest = array[0];
			for (int l = 1; l < array.length; l++) {
				if (array[l].length() < smallest.length()) {
					smallest = array[l];
				}
			}
			cell = row.createCell(2);
			cell.setCellValue(smallest);

			driver.findElement(By.xpath("//*[name()='path' and contains(@d,'M19 6.41L1')]")).click();
			System.out.println("Shortest Option :" + smallest);
			System.out.println("Longest Option :" + longestString);
			
		}
		Thread.sleep(2000);
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		fos.close();
		driver.close();

	}
}
