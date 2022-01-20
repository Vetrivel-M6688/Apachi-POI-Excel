package excelOperations;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UsingIteratorInExcel {

	public static void main(String[] args) throws Exception {

		iteretorCell();
	}

	public static void iteretorCell() throws Exception {
		File fileLoc = new File("./data/test-data_Read.xlsx");
		FileInputStream inputStream = new FileInputStream(fileLoc);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet= workbook.getSheetAt(0);
		
		Iterator<Row> iterate_Row = sheet.iterator();
		while(iterate_Row.hasNext()) {
			Row row = iterate_Row.next();
			Iterator<Cell> iterator_Cell = row.cellIterator();
			while(iterator_Cell.hasNext()) {
				Cell cell = iterator_Cell.next();
				DataFormatter formatter = new DataFormatter();
				String value=formatter.formatCellValue(cell);
				System.out.print(value+"   ");
			}
			System.out.println();
		}
		workbook.close();
	}
}
