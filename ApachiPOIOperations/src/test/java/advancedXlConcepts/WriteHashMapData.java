package advancedXlConcepts;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteHashMapData {

	public static void main(String[] args) throws IOException {

		HashMap<Integer, String> data = new HashMap<Integer, String>();
		data.put(1, "Vetri");
		data.put(5, "Vg");
		data.put(4, "Naren");
		data.put(3, "Prem");
		data.put(2, "Vijay");
		
		String fileLoc = ".//AdavancedDataFiles//RankRecord.xlsx";
		WriteHashMapData wh = new WriteHashMapData();
		wh.hashMapRead(data, fileLoc);
	}

	private void hashMapRead(HashMap<Integer, String> data, String fileLoc) throws IOException {
	
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Ranks");
		
		int rowNum =0;
		for(Map.Entry<Integer, String> entries:data.entrySet()) {
			XSSFRow row = sheet.createRow(rowNum++);
			
			row.createCell(0).setCellValue(entries.getKey());
			row.createCell(1).setCellValue(entries.getValue());
		}
		
		FileOutputStream outputStream = new FileOutputStream(fileLoc);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
		System.out.println("File Written");
	}

}
