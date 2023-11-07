package com.chicmic.JExcel2Pdf.gen;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;




public class GenApplication {

	private static final String FILE_NAME = "For Bank (August).xlsx";
	private static final String rootDirectory = System.getProperty("user.dir");
	public static final String invoiceDirectoriesPath = rootDirectory + "/invoices";
	public static final String documentName = "L1 Request letter for Submission of Export doc.docx";

	public static void main(String[] args) {
		try {
			File excelFile = new File(rootDirectory + "/" +FILE_NAME); // Excel File Read in current Directory
			ExcelSorter excelSorter = new ExcelSorter();
			File sortedExcelFile = excelSorter.excelManager(excelFile); // Excel File Sort acc to column D then F

			ExcelPerformOperations excelPerformOperations = new ExcelPerformOperations();
			excelPerformOperations.excelPerformOperations(sortedExcelFile, rootDirectory);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}


}
