package GainExpertise;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.SortedMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	static String filePath = "C:\\Selenium\\Test Upload.xlsx";
	static FileInputStream file;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet, ActualScenarios, tempTable;
	static Row row;
	static int rowCount, cellCount;
	static HashMap<String, Integer> startPos = new HashMap<>();
	static HashMap<String, Integer> endPos = new HashMap<>();
	static HashMap<Integer, Integer> TempTable = new HashMap<>();
	static HashMap<String, String> testIndex = new HashMap<>();
	static HashMap<Integer, Integer> actualScenarios = new HashMap<>();
	static HashMap<String, String> tempVariables = new HashMap<>();

	static String newScenario;
	static String Subject, TestName, Type, testDesc, Mapping, StepName, Description, ExpectedResult;
	static int startPosition, endPosition;

	public static String[][] fetchSheet(String sheetName) throws Exception {
		System.out.println("Fetching the sheet " + sheetName);
		file = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(file);
		sheet = workbook.getSheet(sheetName);
		rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		row = sheet.getRow(0);
		cellCount = row.getPhysicalNumberOfCells();
		System.out.println("Row count-> " + rowCount + " Column count-> " + cellCount);
		String[][] cellData = new String[rowCount + 1][cellCount];

		for (int i = 0; i <= rowCount; i++) {
			for (int j = 0; j < cellCount; j++) {
				cellData[i][j] = sheet.getRow(i).getCell(j).getStringCellValue();
				// System.out.print(cellData[i][j]+" :: ");
				if (sheetName.equals("TempTable") && j == 0) {

					if (sheet.getRow(i).getCell(4).getStringCellValue().length() > 1) {
						newScenario = sheet.getRow(i).getCell(4).getStringCellValue();
						startPos.put(sheet.getRow(i).getCell(4).getStringCellValue(), i);
					} else {
						// System.out.println(newScenario);
						endPos.put(newScenario, i);
					}

				}
				// return cellData;
			}
			// System.out.println(" ");
		}
		System.out.println(startPos + " \n " + endPos);
		file.close();

		return cellData;
	}

	public static void getScenarioDetails(String Scenario) throws Exception {
		// System.out.println("Getting scenario details for "+Scenario);
		for (String startKey : startPos.keySet()) {
			if (Scenario.equals(startKey)) {
				startPosition = startPos.get(startKey);
			}
		}
		for (String endKey : endPos.keySet()) {
			if (Scenario.equals(endKey)) {
				endPosition = endPos.get(endKey);
			}
		}
		// TempTable.put(startPosition, endPosition);
		System.out.print(startPosition + " " + endPosition + "\n");
		// System.out.println(tempTable);
		writeActualScenarios(Scenario, startPosition, endPosition);
	}

	public static void writeActualScenarios(String Scenario, int startPosition, int endPosition) throws Exception {
		try {
			FileInputStream fis = new FileInputStream("C:\\Selenium\\Test Upload.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet TempTable, ActualScenarios;
			TempTable = workbook.getSheet("TempTable");
			ActualScenarios = workbook.getSheet("Actual Scenarios");
			int rowCountForActualScenarios = ActualScenarios.getLastRowNum() - ActualScenarios.getFirstRowNum() + 1;
			int rowCountForTempTable = TempTable.getLastRowNum() - TempTable.getFirstRowNum();

			for (int l = startPosition; l <= endPosition; l++) {
				Subject = TempTable.getRow(l).getCell(0).getStringCellValue();
				TestName = TempTable.getRow(l).getCell(1).getStringCellValue();
				Type = TempTable.getRow(l).getCell(2).getStringCellValue();
				testDesc = TempTable.getRow(l).getCell(3).getStringCellValue();

				Mapping = TempTable.getRow(l).getCell(4).getStringCellValue();
				StepName = TempTable.getRow(l).getCell(5).getStringCellValue();
				Description = TempTable.getRow(l).getCell(6).getStringCellValue();
				ExpectedResult = TempTable.getRow(l).getCell(7).getStringCellValue();
				copyLines(ActualScenarios, Subject, TestName, Type, testDesc, Mapping, StepName, Description,
						ExpectedResult);

			}

		} catch (IOException e) {
			System.out.println("File not found");
		}

	}

	public static void copyLines(XSSFSheet sheetName, String Subject, String TestName, String Type, String testDesc,
			String Mapping, String StepName, String Description, String ExpectedResult) throws IOException {
		file = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(file);
		int lastRow;
		sheetName = workbook.getSheet("Actual Scenarios");
		// message=message.contains(key)?message.replace(key,
		// replacedText.get(key)):message;

		lastRow = sheetName.getLastRowNum() + 1;
		for (String key : tempVariables.keySet()) {
			if (Subject.contains(key)) {
				Subject = Subject.replace(key, tempVariables.get(key));
			} else if (Mapping.contains(key)) {
				Mapping = Mapping.replace(key, tempVariables.get(key));
			} else if (StepName.contains(key)) {
				StepName = StepName.replace(key, tempVariables.get(key));
			}  else if (ExpectedResult.contains(key)) {
				ExpectedResult = ExpectedResult.replace(key, tempVariables.get(key));
			}
			else if(TestName.contains(key)){
				TestName=TestName.replace(key, tempVariables.get(key));
			}
			
			if (Description.contains(key)) {
				Description = Description.replace(key, tempVariables.get(key));
			}
		}

		sheetName.createRow(lastRow).createCell(0).setCellValue(Subject);
		sheetName.getRow(lastRow).createCell(1).setCellValue(TestName);
		sheetName.getRow(lastRow).createCell(2).setCellValue(Type);
		sheetName.getRow(lastRow).createCell(3).setCellValue(testDesc);
		sheetName.getRow(lastRow).createCell(4).setCellValue(Mapping);
		sheetName.getRow(lastRow).createCell(5).setCellValue(StepName);
		sheetName.getRow(lastRow).createCell(6).setCellValue(Description);
		sheetName.getRow(lastRow).createCell(7).setCellValue(ExpectedResult);
		FileOutputStream fos = new FileOutputStream(filePath);

		workbook.write(fos);
		workbook.close();
	}

	public void scenariosToPick(String sheetName) throws Exception {
		try {
			System.out.println("Fetching the sheet " + sheetName);

			file = new FileInputStream(filePath);
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheet(sheetName);
			rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
			row = sheet.getRow(0);
			cellCount = row.getPhysicalNumberOfCells();
			System.out.println("Row count-> " + rowCount + " Column count-> " + cellCount);

			// get all variable entries as temp array
			String[] temp = new String[cellCount];
			for (int a = 0; a < cellCount; a++) {
				temp[a] = sheet.getRow(0).getCell(a).getStringCellValue();
			}

			for (int i = 1; i <= rowCount; i++) {

				testIndex.put(temp[0], sheet.getRow(i).getCell(0).getStringCellValue());
				testIndex.put(temp[1], sheet.getRow(i).getCell(1).getStringCellValue());
				testIndex.put(temp[2], sheet.getRow(i).getCell(2).getStringCellValue());
				testIndex.put(temp[3], sheet.getRow(i).getCell(3).getStringCellValue());
				testIndex.put(temp[4], sheet.getRow(i).getCell(4).getStringCellValue());
				testIndex.put(temp[5], sheet.getRow(i).getCell(5).getStringCellValue());
				testIndex.put(temp[6], sheet.getRow(i).getCell(6).getStringCellValue());
				testIndex.put(temp[7], sheet.getRow(i).getCell(7).getStringCellValue());
				testIndex.put(temp[8], sheet.getRow(i).getCell(8).getStringCellValue());
				testIndex.put(temp[9], sheet.getRow(i).getCell(9).getStringCellValue());
				testIndex.put(temp[10], sheet.getRow(i).getCell(10).getStringCellValue());
				System.out.println(testIndex);

				// This thread will replace all variables with desired values
				System.out.println("Temporary key value pairs are ");
				for (String key : testIndex.keySet()) {
					if (key.contains("<<")) {
						System.out.println(key + " " + testIndex.get(key));
						tempVariables.put(key, testIndex.get(key));
					}

				}

				// This thread will get all scenarios marked with Y flag & copy
				// it into Actual Scenarios sheet
				for (String key : testIndex.keySet()) {
					if (testIndex.get(key).equals("Y")) {
						System.out.print(key + " ");
						getScenarioDetails(key);
					}
				}

			}

			file.close();
		} catch (FileNotFoundException e) {
			System.out.println("File not found at specified path " + filePath);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws Exception {
		file = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(file);

		ActualScenarios = workbook.getSheet("Actual Scenarios");
		ExcelReader er = new ExcelReader();
		er.fetchSheet("TempTable");
		// getScenarioDetails("FaxCoverPage");
		// getScenarioDetails("PrintLetter");
		// getScenarioDetails("FaxLetter");
		er.scenariosToPick("Test Index");
		// copyLines(ActualScenarios, "FaxLetter","Step 1", "1. Open the
		// <ClientName> Patient Plus.", "Verify the user is able to login to the
		// application successfully.");
		// copyLines(ActualScenarios, "PrintLetter","Step 2", "2. Open the
		// <ClientName> Patient Plus.", "Verify the user is able to login to the
		// application successfully.");
		System.out.println("Write Successful");
		// er.fetchSheet("Test Index");
	}
}
