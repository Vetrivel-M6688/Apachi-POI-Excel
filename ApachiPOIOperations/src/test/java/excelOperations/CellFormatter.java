package excelOperations;
/*
 * using argument passing over the method
 * get the any type of data present in the cell
 * used for loop
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellFormatter {

	public static void main(String[] args) throws IOException {
		String path = "./data/test-data_Read.xlsx";
		String sheetname = "Sheet1";
		readCellFormat(path, sheetname);
	}

	public static void readCellFormat(String filePath, String sheetName) throws IOException {
		File fileLoc = new File(filePath);
		FileInputStream fis = new FileInputStream(fileLoc);

		XSSFWorkbook xlsxWB = new XSSFWorkbook(fis);
		XSSFSheet sheet = xlsxWB.getSheet(sheetName);

		int noOfRows = sheet.getLastRowNum();
		int noOfCols = sheet.getRow(0).getLastCellNum();

		for (int r = 0; r <= noOfRows; r++) {
			XSSFRow row = sheet.getRow(r);
			for (int c = 0; c <= noOfCols; c++) {
				XSSFCell cell = row.getCell(c);
				DataFormatter dataFormatter = new DataFormatter();
				String value = dataFormatter.formatCellValue(cell);
				System.out.print(value + "  ");
			}
			System.out.println();
		}
		xlsxWB.close();
	}
}
