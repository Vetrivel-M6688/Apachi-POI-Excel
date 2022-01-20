package excelOperations;
/*
 * Basic:
 * =================================
 * with and without FileInputStream
 * hard coded no of Rows and Columns
 * Storing file location in String
 * getting the value from the cell only if it is String
 * for loop is used
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleExcelReading {

	public static void main(String[] args) throws IOException {

		SimpleExcelReading ser = new SimpleExcelReading();
		// ser.withoutFIS();
		ser.withFIS();
	}

	public void withoutFIS() throws IOException {
		// getting the file location and stores it in the String
		String fileLocation = "C:\\Users\\Welcome\\Desktop\\MyPrepDocs\\test-data.xlsx";

		// getting the workbook from the fileLocation
		XSSFWorkbook wBook = new XSSFWorkbook(fileLocation);

		// getting the sheet with index 0, means the first sheet of the workbook
		XSSFSheet sheetAt = wBook.getSheetAt(0);

		// looping the 2 rows that we given in the excel, if adding or deleting any rows
		// it won't work

		for (int i = 1; i <= 2; i++) {
			XSSFRow row = sheetAt.getRow(i); // going to the 1st row
			// looping the 2 columns that we given in the excel, if adding or deleting any
			// columns it won't work
			for (int j = 0; j < 2; j++) {
				XSSFCell cell = row.getCell(j); // going to the 0th column
				String stringCellValue = cell.getStringCellValue(); // getting the value from the cell,(Only String
				// values readable)
				System.out.println(stringCellValue);
			}
		}
		wBook.close();
	}

	public void withFIS() throws IOException {

		File fileLoc = new File("C:\\Users\\Welcome\\Desktop\\MyPrepDocs\\test-data.xlsx");

		FileInputStream fis = new FileInputStream(fileLoc);

		XSSFWorkbook workb = new XSSFWorkbook(fis);

		// going to the 1st sheet using the name of the sheet
		XSSFSheet sheet1 = workb.getSheet("Credentials");

		for (int i = 1; i < 3; i++) {
			XSSFRow row = sheet1.getRow(i);
			for (int j = 0; j < 2; j++) {
				XSSFCell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				String CellType = cellType.toString();
				if (CellType.equalsIgnoreCase("String")) {
					System.out.println("This Cell contains String value as: " + cell.getStringCellValue());
				} else if (CellType.equalsIgnoreCase("Numeric")) {
					System.out.println("This Cell contains Numeric value as: " + cell.getNumericCellValue());
				}

			}
			workb.close();
		}

		/*
		 * XSSFRow row = sheet1.getRow(1); XSSFCell cell = row.getCell(0);
		 * 
		 * CellType cellType = cell.getCellType();
		 * System.out.println("Type of the value present in the cell: "+cellType);
		 * String cellTypeString = cellType.toString();
		 * System.out.println("Type of the value present in the cell: "+cellTypeString);
		 * 
		 * String stringCellValue = cell.getStringCellValue();
		 * System.out.println(stringCellValue);
		 */

	}

}
