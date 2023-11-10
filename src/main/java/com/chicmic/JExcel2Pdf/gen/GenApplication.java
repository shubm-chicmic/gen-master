package com.chicmic.JExcel2Pdf.gen;

import java.io.File;
import java.io.IOException;

public class GenApplication {

	public static final boolean autoDeleteFolder = true;
	public static final String FINWNumber = "003FINW232480063"; // Inward remittance Reference Number
	public static final String FILE_NAME = "For Bank (September).xlsx";
	public static final String rootDirectory = System.getProperty("user.dir");
	public static final String invoiceDirectoriesPath = rootDirectory + "/invoices";
	public static final String documentName = "L1 Request letter for Submission of Export doc.docx";
    public static final String fourPointDeclarationDocumentName = "FOUR POINT DECLARATION.docx";
	public static final String fourPointDeclarationDocumentPath = rootDirectory + "/" + fourPointDeclarationDocumentName;

	private static final FolderOperations folderOperations = new FolderOperations();
	public static void main(String[] args) {
		try {
			File excelFile = new File(rootDirectory + "/" +FILE_NAME); // Excel File Read in current Directory
			ExcelSorter excelSorter = new ExcelSorter();
			File sortedExcelFile = excelSorter.excelManager(excelFile); // Excel File Sort acc to column D then F

			ExcelPerformOperations excelPerformOperations = new ExcelPerformOperations();
			excelPerformOperations.excelPerformOperations(sortedExcelFile);

			FourPointDeclaration fourPointDeclaration = new FourPointDeclaration();
			fourPointDeclaration.generateDocument(excelFile);

			// Delete Temp Files
			folderOperations.deleteFile(rootDirectory + "/sorted_" + FILE_NAME);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}


}
