package com.chicmic.JExcel2Pdf.gen;



import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import static com.chicmic.JExcel2Pdf.gen.DateOperations.findMaximumDate;
import static com.chicmic.JExcel2Pdf.gen.DateOperations.getTodaysDate;
import static com.chicmic.JExcel2Pdf.gen.FolderOperations.pathBefore;

public class ExcelPerformOperations {
    String templateDocumentName = "/" + GenApplication.documentName;
    Integer indexOfRecipientColumnD = 3;
    Integer indexOfSOFTEXNumberColumnF = 5;
    // HashMap of updating text , pair < paraindex , runIndex> for updating document with specific text at location paragaraph index and run index
    private HashMap<Pair<Integer, Integer>, Pair<String, String>> textParaRunIndexHashMap = new HashMap<>();

    double billAmount = 0.0; // Initialize the billAmount column g
    double chargesAmount = 0.0; // Initialize the charges column h
    double finalBillAmount = 0.0; // Initialize the final bill column i
    String currentDate = getTodaysDate();

    // Document places changes paraIndex and runindex pair instances
    Pair<Integer, Integer> billAmountIndex = new Pair<>(9, 3);
    Pair<Integer, Integer> chargesAmountIndex = new Pair<>(10, 3);
    Pair<Integer, Integer> finalBillAmountIndex = new Pair<>(11, 3);
    Pair<Integer, Integer> currentDateIndex = new Pair<>(1, 2);
    Pair<Integer, Integer> invoiceDateIndex = new Pair<>(4, 1); // inside table its row and col index
    Pair<Integer, Integer> softexNumberIndex = new Pair<>(3, 1); // inside table its row and col index
    Pair<Integer, Integer> nameOfBuyerIndex = new Pair<>(2, 1); // inside table its row and col index

    FolderOperations folderOperations = new FolderOperations();
    DocxFileOperations docxFileOperations = new DocxFileOperations();
    String excelFilePath = GenApplication.rootDirectory;
    String prevD = "";
    String prevF = "";
    String prevB = ""; // for invoice
    String invoiceDate = "";
    String currentWorkingDirectory = "";

    public void updateDocument() throws IOException {
        billAmount = Math.round(billAmount * 100.0) / 100.0;
        chargesAmount = Math.round(chargesAmount * 100.0) / 100.0;
        finalBillAmount = Math.round(finalBillAmount * 100.0) / 100.0;

        textParaRunIndexHashMap.put(billAmountIndex, new Pair<>(String.valueOf(billAmount), "text"));
        textParaRunIndexHashMap.put(chargesAmountIndex, new Pair<>(String.valueOf(chargesAmount), "text"));
        textParaRunIndexHashMap.put(finalBillAmountIndex, new Pair<>(String.valueOf(finalBillAmount), "text"));
        textParaRunIndexHashMap.put(invoiceDateIndex, new Pair<>(invoiceDate, "table"));
        textParaRunIndexHashMap.put(softexNumberIndex, new Pair<>(prevF, "table"));
        textParaRunIndexHashMap.put(nameOfBuyerIndex, new Pair<>(prevD, "table"));
        docxFileOperations.updateTextAtPosition(excelFilePath + templateDocumentName , currentWorkingDirectory + templateDocumentName, textParaRunIndexHashMap);
    }

    public void excelPerformOperations(File excelFile) throws IOException {
        textParaRunIndexHashMap.put(currentDateIndex, new Pair<>(currentDate, "text"));
        textParaRunIndexHashMap.put(new Pair<>(31, 1), new Pair<>(currentDate, "text"));

        String resultantFilePath = folderOperations.createFolder("Annexure 1", excelFilePath);
        if (resultantFilePath == null) {
            System.out.println("Returning");
            return;
        }

        FileInputStream fis = new FileInputStream(excelFile);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        List<Row> rows = new ArrayList<>();
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            rows.add(row);
        }

        // use this to find the document paragraph index and run index and table row and col index by entering path of doc file
        //docxFileOperations.getParagraphAndRunIndices(excelFilePath);

        int heirarchyIndex = 0;
        currentWorkingDirectory = resultantFilePath;
        for (int i = 0; i < rows.size(); i++) {
            invoiceDate = findMaximumDate(prevB, invoiceDate);

            Row sortedRow = rows.get(i);
            Cell currentCellA = sortedRow.getCell(0);
            Cell currentCellB = sortedRow.getCell(1);
            Cell currentCellC = sortedRow.getCell(2);
            Cell currentCellD = sortedRow.getCell(indexOfRecipientColumnD);
            Cell currentCellF = sortedRow.getCell(indexOfRecipientColumnD + 2);
            String currentD = currentCellD.toString();
            String currentF = currentCellF.toString();
            String currentB = currentCellB.toString();

            double cellValBillAmount = Double.parseDouble(sortedRow.getCell(indexOfRecipientColumnD + 3).toString());
            double cellValChargesAmount = Double.parseDouble(sortedRow.getCell(indexOfRecipientColumnD + 4).toString());
            double cellValFinalBillAmount = Double.parseDouble(sortedRow.getCell(indexOfRecipientColumnD + 5).toString());
//            System.out.println("\u001B[33m amount g" + i + " : cellValBillAmount = " + cellValBillAmount + " cellValChargesAmount = " + cellValChargesAmount + " cellValFinalBillAmount : " + cellValFinalBillAmount +"\u001B[0m");

            if (currentCellD != null) {
                if (prevD.equals(currentD)) {
                    if (currentF.equals(prevF)) {
                        cellValBillAmount = Math.round(cellValBillAmount * 100.0) / 100.0;
                        cellValChargesAmount = Math.round(cellValChargesAmount * 100.0) / 100.0;
                        cellValFinalBillAmount = Math.round(cellValFinalBillAmount * 100.0) / 100.0;
                        billAmount += cellValBillAmount;
                        chargesAmount += cellValChargesAmount;
                        finalBillAmount += cellValFinalBillAmount;
                    } else {
                        updateDocument();
                        heirarchyIndex++;
                        invoiceDate = "";
                        currentWorkingDirectory = folderOperations.createFolder(String.valueOf(heirarchyIndex), pathBefore(currentWorkingDirectory)); // create folder with name = '1'
                        billAmount = cellValBillAmount;
                        chargesAmount = cellValChargesAmount;
                        finalBillAmount = cellValFinalBillAmount;
                    }
                } else {
                    updateDocument();
                    String path = folderOperations.createFolder(currentD, resultantFilePath);
                    heirarchyIndex = 1;
                    currentWorkingDirectory = folderOperations.createFolder(String.valueOf(heirarchyIndex), path); // create folder with name = '1'
                    invoiceDate = "";
                    billAmount = cellValBillAmount;
                    chargesAmount = cellValChargesAmount;
                    finalBillAmount = cellValFinalBillAmount;
                }
                updateDocument();
                // Invoice pdf search and save in current working directory
                File invoiceFile = folderOperations.searchForFile(GenApplication.invoiceDirectoriesPath, currentCellC.toString() + ".pdf");
                folderOperations.saveFileToOutputPath(invoiceFile, currentWorkingDirectory);

                prevD = currentD;
                prevF = currentF;
                prevB = String.valueOf(currentB);
            }

        }

        System.out.println("\u001B[35m" + getClass().getName() + " : Operation Completed without exception !\u001B[0m" );
    }
}


