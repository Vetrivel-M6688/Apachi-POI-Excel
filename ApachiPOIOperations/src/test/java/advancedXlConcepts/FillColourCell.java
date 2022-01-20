package advancedXlConcepts;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FillColourCell {

	public static void main(String[] args) throws IOException {
		FillColourCell fcc = new FillColourCell();
		fcc.colorCell();
	}

	public void colorCell() throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet= workbook.createSheet("ColorTest");
		
		XSSFRow row=sheet.createRow(3);
		
		XSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
		cellStyle.setFillPattern(FillPatternType.BRICKS);
		
		XSSFCell cell = row.createCell(5);
		cell.setCellValue("Background");
		cell.setCellStyle(cellStyle);
		
		cellStyle = workbook.createCellStyle();
		cellStyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell = row.createCell(6);
		cell.setCellValue("ForeGround");
		cell.setCellStyle(cellStyle);
		
		FileOutputStream outputStream = new FileOutputStream(".//AdavancedDataFiles//ColorExcel.xlsx");
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
		System.out.println("Job Done");
	}
}
