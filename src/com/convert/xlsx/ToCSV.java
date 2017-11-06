package com.convert.xlsx;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.Iterator;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ToCSV {
	public static void main(String[] args) throws IOException {

		//InputStream ExcelFileToRead = new FileInputStream("C:/Users/Yehua/Desktop/project_release/All-test-final-with-location-code.xlsx");
		InputStream ExcelFileToRead = new FileInputStream("/Users/KaikaiFu/Desktop/excel/scientist.xlsx");
		StringBuilder sb = new StringBuilder();
		
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

		XSSFWorkbook test = new XSSFWorkbook();

		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow row;
		XSSFCell cell;

		Iterator rows = sheet.rowIterator();
		
		while (rows.hasNext()) {
			row = (XSSFRow) rows.next();
			Iterator cells = row.cellIterator();
			while (cells.hasNext()) {
				cell = (XSSFCell) cells.next();

				if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {				
						sb.append(cell.getStringCellValue()+"|");
					System.out.print(cell.getStringCellValue() + " ");
				} else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
					String columnName = CellReference.convertNumToColString(cell.getColumnIndex()); 
					if(columnName.equalsIgnoreCase("L") || columnName.equalsIgnoreCase("M")){
						sb.append(cell.getNumericCellValue() + "|");
					} else {
						sb.append((int)cell.getNumericCellValue() + "|");
					}
					//sb.append("1.11"+"|");
					System.out.print(cell.getNumericCellValue() + " ");
				} else {
					// U Can Handel Boolean, Formula, Errors
				}
			}
			sb.append("\n");
			
			
			System.out.println();
		}
		
		try{
		    
		      OutputStreamWriter pw = new OutputStreamWriter(new FileOutputStream("/Users/KaikaiFu/Desktop/excel/downloaded.csv"),"UTF-8");
		      
		      byte b[] = {(byte)0xEF,(byte)0xBB,(byte)0xBF};
		      String bom = new String(b);
		      //pw.write(bom);
		      pw.write(sb.toString());
		      
		      pw.close();
		      System.out.println("Done");

		     }catch(IOException e){
		      e.printStackTrace();
		     }
	}

}
