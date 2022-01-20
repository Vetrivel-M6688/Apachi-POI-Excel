package advancedXlConcepts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadHashMapData {

	public static void main(String[] args) throws IOException {

		ReadHashMapData rh = new ReadHashMapData();
		rh.readHashMap();

	}

	public void readHashMap() throws IOException {

		HashMap<Integer, String> data = new HashMap<Integer, String>();
		FileInputStream inputStream = new FileInputStream(".//AdavancedDataFiles//RankRecord.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

		XSSFSheet sheet = workbook.getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();

		// Reading data from Excel
		for (int r = 0; r <= lastRowNum; r++) {
			double keys = sheet.getRow(r).getCell(0).getNumericCellValue();
			String values = sheet.getRow(r).getCell(1).getStringCellValue();

			data.put((int) keys, values);
		}

		// Reading data from HashMap
		for (Map.Entry<Integer, String> entries : data.entrySet()) {
			System.out.println("Rank: " + entries.getKey() + "  Name: " + entries.getValue());
		}
	}

}
