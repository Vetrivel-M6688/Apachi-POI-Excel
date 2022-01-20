package excelOperations;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteXLwithArray {

	public static void main(String[] args) throws IOException {

		withForLoop();
		withForEachLoop();
	}

	static XSSFWorkbook 	empDeatails = new XSSFWorkbook();
	static File fileLoc = new File("./data/EmpDet_Array.xlsx"); 

	public static void withForLoop() throws IOException {

		Object[][] data = { { "EmpId", "Name", "Salary" }, 
				{ 111, "Vetri", 30000 }, 
				{ 222, "Vijay", 31000 },
				{ 333, "Naren", 41000 }, 
				{ 444, "VG", 45000 }, 
				{ 555, "Prem", 33000 } };
		XSSFSheet sheet = empDeatails.createSheet("SalaryDetails");

		int noOfRows = data.length; // will check the tot rows of data present in the array
		int noOfCols = data[0].length; // will check the tot cols of data present at the 0th ROW in array
		System.out.println("Total no of Rows: " + noOfRows);
		System.out.println("Total no of Columns at 0th Row: " + noOfCols);

		for (int r = 0; r < noOfRows; r++) {
			XSSFRow row = sheet.createRow(r);
			for (int c = 0; c < noOfCols; c++) {
				XSSFCell cell = row.createCell(c);
				Object value = data[r][c];

				if (value instanceof String) {
					String stringVal = value.toString();
					cell.setCellValue(stringVal);
				} else if (value instanceof Integer) {
					cell.setCellValue((Integer) value);
				} else if (value instanceof Boolean) {
					cell.setCellValue((Boolean) value);
				}
			}
		}

		FileOutputStream outputStream = new FileOutputStream(fileLoc);
		empDeatails.write(outputStream);
		outputStream.close();
		//empDeatails.close();
		System.out.println("File written successfully............");
	}
	public static void withForEachLoop() throws IOException {
		
		XSSFSheet sheet = empDeatails.createSheet("EmpDetForeach");
		Object[][] data = {	{"S.No","Name","Time"},
									{1,"Vetri",1.30f},
									{2,"Vijay",10.20f},
									{3,"VG",3.00f},
									{4,"Prem",5.45f},
									{5,"Naren",6.00f}
									};
		
		int rowCount = 0;
		for (Object[] rows : data) {
			XSSFRow row = sheet.createRow(rowCount++);
			int columnCount =0;
			for(Object cellValues : rows) {
				XSSFCell cell = row.createCell(columnCount++);
				if(cellValues instanceof String) {
					cell.setCellValue((String)cellValues);
				}else if(cellValues instanceof Integer) {
					cell.setCellValue((Integer)cellValues);
				}else if(cellValues instanceof Float)
					cell.setCellValue((Float)cellValues);
			}
		}

		FileOutputStream outputStream = new FileOutputStream(fileLoc);
		empDeatails.write(outputStream);
		empDeatails.close();
		outputStream.close();
		System.out.println("File Written Successflly!!!!");
	}
}
