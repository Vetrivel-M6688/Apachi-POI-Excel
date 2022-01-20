package excelOperations;
/*
 * With FIS and File concept for storing the file location
 * without hard coded, loop keep running until the last Row and the Columns
 * get the value from the cell only if it has a String value
 * for loop is used
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dynamicRowCol {

	public static void main(String[] args) throws IOException {
		dynamic();
	}

	public static void dynamic() throws IOException {

		FileInputStream fis = new FileInputStream(new File("./data/test-data_Read.xlsx"));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheetAt = wb.getSheetAt(0);

		// getLastRowNum() is used to get the last updated row in the workbook, whether
		// it might add or delete the value in the row
		int totRowsExcHeader = sheetAt.getLastRowNum();
		System.out.println("Total no of rows without Headers: " + totRowsExcHeader);
		int totRowsIncHeader = sheetAt.getPhysicalNumberOfRows();
		System.out.println("Total no of rows with Headers: " + totRowsIncHeader);

		// getLastCellNum() is used to get the last updated cell in the workbook,
		// whether it might add or delete the value
		short lastCellNum = sheetAt.getRow(0).getLastCellNum();
		System.out.println("last Cell Number: " + lastCellNum);

		for (int i = 1; i <= totRowsExcHeader; i++) {
			XSSFRow row = sheetAt.getRow(i);
			for (int j = 0; j < lastCellNum; j++) {
				XSSFCell cell = row.getCell(j);
				String stringCellValue = cell.getStringCellValue();  // the problem here is we can either get String or Numeric value from the cell
				System.out.println(stringCellValue);
			}
		}
		wb.close();
	}

}
