package excelOperations;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteXlwithArrayList {

	public static void main(String[] args) {
		try {
			usingArrayList();
		} catch (IOException e) {
			System.out.println("Something went wrong!!, Please check the file creation operation again");
			e.printStackTrace();
		}
	}

	public static void usingArrayList() throws IOException {

		ArrayList<Object[]> data = new ArrayList<Object[]>();
		data.add(new Object[] { "S.No", "BrandName", "Amount" });
		data.add(new Object[] { 1, "Xiomi", 15000 });
		data.add(new Object[] { 2, "Samsung", 40000 });
		data.add(new Object[] { 3, "Vivo", 13000 });
		data.add(new Object[] { 4, "Apple", 150000 });

		File filePath = new File("./data/MobileDet_ArrayList.xlsx");
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet("MobileBrands");

		int rowNum = 0;
		for (Object[] rows : data) {
			XSSFRow row = sheet.createRow(rowNum++);
			int colNum = 0;
			for (Object values : rows) {
				XSSFCell cell = row.createCell(colNum++);
				if (values instanceof String) {
					cell.setCellValue((String) values);
				} else if (values instanceof Integer) {
					cell.setCellValue((Integer) values);
				} else {
					System.out.println("Datatype Not matched");
				}
			}
		}

		FileOutputStream outStream = new FileOutputStream(filePath);
		workBook.write(outStream);
		outStream.close();
		workBook.close();
		System.out.println("File Successfully updted.....");
	}

}
