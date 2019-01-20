import java.util.Map;

public class RunApp {

	public static void main(String[] args) throws Exception {

		String BSFilePath = "d:\\hc\\CalculatorData.xlsx";
		String BSSheetName = "BLSteps";

		String ORFilePath = "d:\\hc\\CalculatorData.xlsx";
		String ORSheetName = "ObjectRepo";

		Map<String, String> result = SeleniumKeyManager.Execute(BSFilePath, BSSheetName, ORFilePath, ORSheetName);
		System.out.println(result);

	}

}
