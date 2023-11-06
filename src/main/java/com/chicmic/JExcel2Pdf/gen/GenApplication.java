package com.chicmic.JExcel2Pdf.gen;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;




public class GenApplication {

	private static final String FILE_NAME = "For Bank (August).xlsx";
	private static final String path = System.getProperty("user.dir");

	public static void main(String[] args) {
		try {



			File excelFile = new File(path + "/" +FILE_NAME); // Excel File Read in current Directory
			ExcelSorter excelSorter = new ExcelSorter();
			File sortedExcelFile = excelSorter.excelManager(excelFile); // Excel File Sort acc to column D then F

			ExcelPerformOperations excelPerformOperations = new ExcelPerformOperations();
			excelPerformOperations.excelPerformOperations(sortedExcelFile, path);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}


}
