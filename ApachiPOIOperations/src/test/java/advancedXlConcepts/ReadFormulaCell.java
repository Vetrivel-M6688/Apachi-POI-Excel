package advancedXlConcepts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFormulaCell {

	public static void main(String[] args) {

		String filePath = "./AdavancedDataFiles/ReadFormula.xlsx";
		ReadFormulaCell rfc = new ReadFormulaCell(filePath);
	}

	public ReadFormulaCell(String fileLoc) {

		try {
			FileInputStream stream = new FileInputStream(fileLoc);
			XSSFWorkbook workbook = new XSSFWorkbook(stream);
			XSSFSheet sheet = workbook.getSheetAt(0);
			int rowsCount = sheet.getLastRowNum();
			int colCount = sheet.getRow(0).getLastCellNum();
			for (int r = 0; r < rowsCount; r++) {
				XSSFRow row = sheet.getRow(r);
				for (int c = 0; c < colCount; c++) {
					XSSFCell cell = row.getCell(c);

					switch (cell.getCellType()) {
					case STRING:
						System.out.print(cell.getStringCellValue());
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
						
						// will get the value from the formulated cell
					case FORMULA:
						System.out.print(cell.getNumericCellValue());
						break;
					default:
						System.out.println("No cell type matches");
						break;
					}
					System.out.print("|");
				}
				System.out.println();
			}
			workbook.close();
		} catch (FileNotFoundException e) {
			System.out.println("Mentioned file is not present or may be type is different");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("WorkBook not found");
			e.printStackTrace();
		}
		System.out.println("Read data Successfully!!!");
	}
	
	/* USING THIS WILL GIVE YOU THE FORMULA WHICH HANDLE IN THE CELL (EX: =SUM(A1+A2)
	 * DataFormatter formate = new DataFormatter(); 
	 * String formatCellValue = formate.formatCellValue(cell); 
	 * System.out.print(formatCellValue);
	 */

}
