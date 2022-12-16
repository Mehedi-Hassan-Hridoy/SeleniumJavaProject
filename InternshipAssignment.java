package projects;

import java.io.File;
import java.io.FileInputStream;
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

public class InternshipAssignment {
	public static void main(String[] args) throws InterruptedException, IOException {

		System.setProperty("webdriver.chrome.driver","E://CS//Selenium//Chrome Driver//chromedriver_win32/chromedriver.exe");
		WebDriver driver = new ChromeDriver();// open the browser
		driver.manage().window().maximize();
		driver.get("https://www.google.com");
		driver.findElement(By.xpath("//a[contains(text(),'English')]")).click();

		File file = new File("E:\\CS\\Java\\Java Project\\SeleniumProjects\\Excel\\Excel.xlsx");
		FileInputStream inputstream = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(inputstream);
		String[] sheetArray = new String[] { "Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday","Friday" };
		
		String date = "";
		int indexSheet;
		for (indexSheet = 0; indexSheet < sheetArray.length; indexSheet++)
		{
		
			date = sheetArray[indexSheet];
			XSSFSheet sheet = wb.getSheet(date);
			XSSFRow row = null;
			XSSFCell cell = null;
			String keyword = null;
			System.out.println(sheet.getLastRowNum());
			for (int i = 2; i <= sheet.getLastRowNum(); i++) {
				row = sheet.getRow(i);
				for (int j = 2; j <= row.getLastCellNum(); j++) {
					cell = row.getCell(j);
					if (j == 2) {
						keyword = cell.getStringCellValue();
					}
				}
				
				driver.findElement(By.name("q")).sendKeys(keyword);
				Thread.sleep(2000);
				List<WebElement> listOfElements = driver.findElements(By.xpath("//ul/li[@role='presentation']"));
				int elementSize = listOfElements.size();
				String[] optionsArray = new String[elementSize];
				int a = 0;
				int maxLength = 0;
				String longestOptions = "";

				for (WebElement in : listOfElements) {
					optionsArray[a] = in.getText();
					a++;
				}

				for (String s : optionsArray) {
					if (s.length() > maxLength) {
						maxLength = s.length();
						longestOptions = s;
					}
				}
				cell = row.createCell(3);
				cell.setCellValue(longestOptions);

				String shortestOptions = optionsArray[0];
				for (int l = 1; l < optionsArray.length; l++) {
					if (optionsArray[l].length() < shortestOptions.length()) {
						shortestOptions = optionsArray[l];
					}
				}
				cell = row.createCell(4);
				cell.setCellValue(shortestOptions);

				driver.findElement(By.xpath("//*[name()='path' and contains(@d,'M19 6.41L1')]")).click();
				System.out.println("Shortest Option :" + shortestOptions);
				System.out.println("Longest Option :" + longestOptions);

			}

		}
		Thread.sleep(2000);
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		fos.close();
		driver.close();

	}

}
