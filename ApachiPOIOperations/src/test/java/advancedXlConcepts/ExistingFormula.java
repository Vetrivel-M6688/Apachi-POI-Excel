package advancedXlConcepts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExistingFormula {

	public static void main(String[] args) throws IOException {

		ExistingFormula ef = new ExistingFormula();
		ef.writeTotalFormula();
	}

	public void writeTotalFormula() throws IOException {

		File filepath = new File(".//AdavancedDataFiles//WriteFormula.xlsx");
		FileInputStream inpuStream = new FileInputStream(filepath);
		XSSFWorkbook workbook = new XSSFWorkbook(inpuStream);
		XSSFSheet sheet = workbook.getSheet("Price");

		sheet.getRow(6).getCell(2).setCellFormula("SUM(C2:C6)");

		inpuStream.close();
		FileOutputStream outputStream = new FileOutputStream(filepath);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
		System.out.println("Found the Total");
	}
}
