import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MyDataProvider {

	// https://codepen.io/TrainCode/pen/qmQVRm?editors=1100
	public static Object[][] ReadExcel(String filepath, String sheetName, int column) throws Exception {

		XSSFWorkbook excel = new XSSFWorkbook(filepath);

		XSSFSheet datasheet = excel.getSheet(sheetName);

		int rowCount = datasheet.getLastRowNum();

		Object[][] fillData = new Object[rowCount][column];
		
		int headerCounter = 1;
		for (int r = headerCounter; r <= rowCount; r++) {
			for (int c = 0; c < column; c++) {

				if (datasheet.getRow(r).getCell(c).getCellTypeEnum() == CellType.STRING) {
					String data = datasheet.getRow(r).getCell(c).getStringCellValue();
					fillData[r-headerCounter][c] =data;
					System.out.println(data);
				} else if (datasheet.getRow(r).getCell(c).getCellTypeEnum() == CellType.NUMERIC) {
					double data = datasheet.getRow(r).getCell(c).getNumericCellValue();
					System.out.println(data);
					fillData[r-headerCounter][c] =String.valueOf(data);
				} else {
					System.out.println("Value miss");
				}
			}
		}
		
		return fillData;

	}

}
