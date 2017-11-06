package com.convert.xls;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ToCSV {
public static void main(String[] args) {
	
	
	Workbook readwb = null;
	try {
		InputStream instream = new FileInputStream("D:/附表一：深海关键技术与装备领域主要学者列表.xls");
		 readwb = Workbook.getWorkbook(instream);
		 Sheet readsheet = readwb.getSheet(0); 
		 int rsColumns = readsheet.getColumns();
		 int rsRows = readsheet.getRows();
		 for (int i = 0; i < rsRows; i++)   
			  
         {   

             for (int j = 0; j < rsColumns; j++)   

             {   

                 Cell cell = readsheet.getCell(j, i);   

                 System.out.print(cell.getContents() + " ");   

             }   

             System.out.println();   

         }   
		
	} catch (FileNotFoundException e1) {
		// TODO Auto-generated catch block
		e1.printStackTrace();
	} catch (BiffException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	
	
	String header = "序号|姓名|国家|单位|职称|总发文|总引用|H指数|联系方式|个人主页";
	String content = "1|Andreae, Meinrat O|德国|马克斯·普朗克化学研究所|教授|439|30036|85|biogeo@mpic.de|http://www.mpg.de/443650/chemie_wissM";
	try{
	      String data = header + "\n" +content;

	      File file =new File("javademo.csv");

	      //if file doesnt exists, then create it
	      if(!file.exists()){
	       file.createNewFile();
	      }

	      //true = append file
	      FileWriter fileWritter = new FileWriter(file.getName(),true);
	             BufferedWriter bufferWritter = new BufferedWriter(fileWritter);
	             bufferWritter.write(data);
	             bufferWritter.close();

	         System.out.println("Done");

	     }catch(IOException e){
	      e.printStackTrace();
	     }
	    }
	
}
