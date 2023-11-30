package com.chicmic.engine;

import com.chicmic.ExcelReadAndDataTransfer.ExcelPerformOperations;
import com.chicmic.ExcelReadAndDataTransfer.ExcelSorter;
import com.chicmic.FourPointDeclarationGen.FourPointDeclaration;
import com.chicmic.Util.FolderOperations.FolderOperations;

import java.io.File;
import java.io.IOException;

public class MainRunner {

	public static final boolean autoDeleteFolder = true;
	public static final boolean deleteTempFiles = true;
	public static final String FINWNumber = "003FINW232480063"; // Inward remittance Reference Number
	public static final String FILE_NAME = "For Bank (September).xlsx";
	public static final String FILE_NAME_WITHOUT_EXTENSION = FolderOperations.getFileNameWithoutExtension(FILE_NAME);
	public static final String rootDirectory = System.getProperty("user.dir");
	public static final String invoiceDirectoriesPath = rootDirectory + "/invoices";
	public static final String documentName = "L1 Request letter for Submission of Export doc.docx";
    public static final String fourPointDeclarationDocumentName = "FOUR POINT DECLARATION.docx";
	public static final String fourPointDeclarationDocumentPath = rootDirectory + "/" + fourPointDeclarationDocumentName;

	public static final String exportRegularisationDocumentName = "EXPORT REGULARISATION TEMPLATE/EXPORT REGULARISATION FORMAT.doc";

	public static void main(String[] args) {
		FolderHeirarchyAndFourPoint.main(args);
	}


}
