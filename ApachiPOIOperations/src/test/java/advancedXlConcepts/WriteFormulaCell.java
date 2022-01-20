package advancedXlConcepts;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell {

	public static void main(String[] args) throws IOException {

		WriteFormulaCell wfc = new WriteFormulaCell();
		wfc.writeFormulCell();
	}

	public void writeFormulCell() throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("AdditionOperation");
		// without Loop
		XSSFRow row = sheet.createRow(0);
		
		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		row.createCell(2).setCellValue(30);
		
		row.createCell(3).setCellFormula("A1+B1+C1");
		FileOutputStream outputStream = new FileOutputStream("./AdavancedDataFiles/WriteFormula.xlsx");
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
		System.out.println("File Written Successfully");
	}
	
}
