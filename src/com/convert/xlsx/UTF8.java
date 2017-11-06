package com.convert.xlsx;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;

public class UTF8 {
	public static void main(String[] args) throws IOException {
		InputStream ExcelFileToRead = new FileInputStream("C:/Users/Yehua/Desktop/project_release/All.xlsx");
		InputStreamReader sr = new InputStreamReader(ExcelFileToRead, "utf-8"); 
		BufferedReader reader = new BufferedReader(sr); 
		String line;
		while ((line = reader.readLine()) != null) {
			System.out.println(line);
		}
	}
}
