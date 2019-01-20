import java.lang.annotation.Target;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.formula.functions.XYNumericFunction;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

public class SeleniumKeyManager {

	public static Map<String, String> Execute(
			String BSFilePath,String BSSheetName,
			String ORFilePath,String ORSheetName
			) throws Exception {

		WebDriver driver = null;

		//Object[][] steps = MyDataProvider.ReadExcel("d:\\hc\\CalculatorData.xlsx", "BLSteps", 4);
		Object[][] steps = MyDataProvider.ReadExcel(BSFilePath, BSSheetName, 4);

		// Read ObjectRepo File
		//Object[][] ObjectRepoData = MyDataProvider.ReadExcel("d:\\hc\\CalculatorData.xlsx", "ObjectRepo", 2);
		Object[][] ObjectRepoData = MyDataProvider.ReadExcel(ORFilePath, ORSheetName, 2);

		// Convert ObjectRepo File into searchable collection
		Map<String, String> searchObjectRepo = new HashMap<>();
		for (int r = 0; r < ObjectRepoData.length; r++) {

			String key = (String) ObjectRepoData[r][0];
			String actualXpath = (String) ObjectRepoData[r][1];

			searchObjectRepo.put(key, actualXpath);
		}

		Map<String, String> stepNumberResult = new HashMap<>();

		for (int r = 0; r < steps.length; r++) {

			String command = (String) steps[r][0];

			// String xpathTarget = (String) steps[r][1];
			String targetKey = (String) steps[r][1];
			String xpathTarget = searchObjectRepo.get(targetKey);

			String value = (String) steps[r][2];

			String stepNumber = (String) steps[r][3];

			System.out.println(command + " " + xpathTarget + " " + value);
			if (command.equalsIgnoreCase("OpenChorme")) {
				System.setProperty("webdriver.chrome.driver", value);
				driver = new ChromeDriver();
				driver.get(xpathTarget);
			} else if (command.equalsIgnoreCase("OpenFirefox")) {
				System.setProperty("webdriver.firefox.driver", value);
				driver = new FirefoxDriver();
				driver.get(xpathTarget);
			} else if (command.equalsIgnoreCase("type")) {
				driver.findElement(By.xpath(xpathTarget)).sendKeys(value);
			} else if (command.equalsIgnoreCase("click")) {
				driver.findElement(By.xpath(xpathTarget)).click();
			} else if (command.equalsIgnoreCase("select")) {
				WebElement ele = driver.findElement(By.xpath(xpathTarget));
				Select sel = new Select(ele);
				sel.selectByVisibleText(value);
			} else if (command.equalsIgnoreCase("readValue")) {
				
				WebElement control = driver.findElement(By.xpath(xpathTarget));
				String result = control.getAttribute("value");
				stepNumberResult.put(stepNumber, result);
				
				
				
			} else if (command.equalsIgnoreCase("read")) {
				// input//select//other
				WebElement control = driver.findElement(By.xpath(xpathTarget));
				if (control.getTagName().equalsIgnoreCase("input")) {
					String result = control.getAttribute("value");
					stepNumberResult.put(stepNumber, result);
				} else if (control.getTagName().equalsIgnoreCase("select")) {
					String result = new Select(control).getFirstSelectedOption().getText();
					stepNumberResult.put(stepNumber, result);
				} else {
					String result = control.getText();
					stepNumberResult.put(stepNumber, result);
				}

			} else {
				throw new Exception("Unsupported Command " + command);
			}

		}
		return stepNumberResult;

	}

}
