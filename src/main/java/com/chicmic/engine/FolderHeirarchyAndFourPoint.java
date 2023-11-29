package com.chicmic.engine;

import com.chicmic.ExcelReadAndDataTransfer.ExcelPerformOperations;
import com.chicmic.ExcelReadAndDataTransfer.ExcelSorter;
import com.chicmic.FourPointDeclarationGen.FourPointDeclaration;
import com.chicmic.Util.FolderOperations.FolderOperations;

import java.io.File;
import java.io.IOException;

public class FolderHeirarchyAndFourPoint {
    private static final FolderOperations folderOperations = new FolderOperations();
    public static void main(String[] args) {
        try {
            File excelFile = new File(MainRunner.rootDirectory + "/" +MainRunner.FILE_NAME); // Excel File Read in current Directory
            ExcelSorter excelSorter = new ExcelSorter();
            File sortedExcelFile = excelSorter.excelManager(excelFile); // Excel File Sort acc to column D then F

            ExcelPerformOperations excelPerformOperations = new ExcelPerformOperations();
            excelPerformOperations.excelPerformOperations(sortedExcelFile);

            FourPointDeclaration fourPointDeclaration = new FourPointDeclaration();
            fourPointDeclaration.generateDocument(excelFile);

            // Delete Temp Files
            folderOperations.deleteFile(MainRunner.rootDirectory + "/sorted_" + MainRunner.FILE_NAME);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
