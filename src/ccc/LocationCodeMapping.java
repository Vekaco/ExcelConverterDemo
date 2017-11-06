package ccc;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LocationCodeMapping {
	public static void main(String[] args) throws IOException {
		InputStream is = new FileInputStream("C:/Users/Yehua/Desktop/project_release/All-test-final.xlsx");
		StringBuilder sb = new StringBuilder();
		StringBuilder value = new StringBuilder();
		
		XSSFWorkbook wb = new XSSFWorkbook(is);
		XSSFSheet sheet = wb.getSheetAt(1);
		sb.append("{");
		for(int i = 0; i < sheet.getLastRowNum() +1 ; i++){
			Row row = sheet.getRow(i);
			String name = row.getCell(0).getStringCellValue();
			double longitude = row.getCell(1).getNumericCellValue();
			double latitude = row.getCell(2).getNumericCellValue();
			String code = row.getCell(3).getStringCellValue();
			sb.append(code).append(":{'latitude':").append(longitude).append(",").append("'longitude':").append(latitude).append("},").append("\n");
			value.append("{").append("'code':").append("'").append(code).append("'").append(",");
			value.append("'name':").append("'").append(name).append("'").append(",");
			value.append("'value':").append((int) (Math.random() * 100)).append("}").append(",").append("\n");
		}
		value.deleteCharAt(value.lastIndexOf(","));
		sb.deleteCharAt(sb.lastIndexOf(","));
		sb.append("}");
		System.out.println(sb.toString());
	}
}
