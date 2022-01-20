package advancedXlConcepts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadProtectedXL {

	public static void main(String[] args) {

		String filePath = ".\\AdavancedDataFiles\\PasswordBook.xlsx";
		Scanner scan = new Scanner(System.in);
		System.out.println("\"PasswordBook\" is Protected, Please enter the password:");
		String password = scan.next();
		ReadProtectedXL rp = new ReadProtectedXL();
		rp.readPasswordXL(filePath, password);
		scan.close();
	}

	private void readPasswordXL(String filePath, String password) {

		try {
			FileInputStream inputStream = new FileInputStream(filePath);

			/*
			 * Normal way for getting the password protected workbook SYNTAX: XSSFWorkbook
			 * workbook = (XSSFWorkbook) WorkbookFactory.create(inputStream,password);
			 */

			// 'Workbook' is a interface and 'WorkbookFactory' is class, so we use the
			// workbook like this
			Workbook workbook = WorkbookFactory.create(inputStream, password);
			Sheet sheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = sheet.iterator();
			while (iterator.hasNext()) {
				Row row = iterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					DataFormatter formatter = new DataFormatter();
					String value = formatter.formatCellValue(cell);
					System.out.print(value);
					System.out.print("|");
				}
				System.out.println();
			}
			System.out.println("File Written Successfully");
		} catch (FileNotFoundException e) {
			System.out.println("Given file is not found in the folder");
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			System.out.println("File is Protect, Please provide the valid password");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("Something went wrong");
			e.printStackTrace();
		}

	}

}
