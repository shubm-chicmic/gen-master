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
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.chicmic.JExcel2Pdf.gen.DateConverter.findGreatestDate;
import static com.chicmic.JExcel2Pdf.gen.DateConverter.getTodaysDate;
import static com.chicmic.JExcel2Pdf.gen.FolderOperations.pathBefore;

public class ExcelPerformOperations {
    Integer indexOfRecipientColumnD = 3;
    Integer indexOfSOFTEXNumberColumnF = 5;
    // HashMap of updating text , pair < paraindex , runIndex> for updating document with specific text at location paragaraph index and run index
    private HashMap<Pair<Integer, Integer>, Pair<String, String>> textParaRunIndexHashMap = new HashMap<>();

    double billAmount = 0.0; // Initialize the billAmount column g
    double chargesAmount = 0.0; // Initialize the charges column h
    double finalBillAmount = 0.0; // Initialize the final bill column i
    String currentDate = getTodaysDate();

    // Document places changes paraIndex and runindex pair instances
    Pair billAmountPair = new Pair<>(9, 3);
    Pair chargesAmountPair = new Pair<>(10, 3);
    Pair finalBillAmountPair = new Pair<>(11, 3);
    Pair currentDatePair = new Pair<>(1, 2);
    Pair invoiceDatePair = new Pair<>(4, 1); // inside table its row and col index
    Pair softexNumberPair = new Pair<>(3, 1); // inside table its row and col index
    Pair nameOfBuyerPair = new Pair<>(2, 1); // inside table its row and col index


    public void excelPerformOperations(File excelFile, String rootDirectoryPath) throws IOException {

        textParaRunIndexHashMap.put(currentDatePair, new Pair<>(currentDate, "text"));
        textParaRunIndexHashMap.put(new Pair<>(31, 1), new Pair<>(currentDate, "text"));

        FolderOperations folderOperations = new FolderOperations();
        DocxFileOperations docxFileOperations = new DocxFileOperations();

        String excelFilePath = rootDirectoryPath;
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

        String prevD = "";
        String prevF = "";
        String prevB = ""; // for invoice
        String invoiceDate = "";
        int heirarchyIndex = 0;
        String currentWorkingDirectory = resultantFilePath;
        for (int i = 0; i < rows.size(); i++) {
            invoiceDate = findGreatestDate(prevB, invoiceDate);

            Row sortedRow = rows.get(i);
            Cell currentCellA = sortedRow.getCell(0);
            Cell currentCellB = sortedRow.getCell(1);
            Cell currentCellC = sortedRow.getCell(2);
            Cell currentCellD = sortedRow.getCell(indexOfRecipientColumnD);
            Cell currentCellF = sortedRow.getCell(indexOfRecipientColumnD + 2);
            String currentD = currentCellD.toString();
            String currentF = currentCellF.toString();
            String currentB = currentCellB.toString();
//            System.out.println("\u001B[35m date greatest of index = " + i + " : prevB = " + prevB + " currentB = " + currentB + " invoidce : " + invoiceDate +"\u001B[0m");


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


                        billAmount = Math.round(billAmount * 100.0) / 100.0;
                        chargesAmount = Math.round(chargesAmount * 100.0) / 100.0;
                        finalBillAmount = Math.round(finalBillAmount * 100.0) / 100.0;

                        textParaRunIndexHashMap.put(billAmountPair, new Pair<>(String.valueOf(billAmount), "text"));
                        textParaRunIndexHashMap.put(chargesAmountPair, new Pair<>(String.valueOf(chargesAmount), "text"));
                        textParaRunIndexHashMap.put(finalBillAmountPair, new Pair<>(String.valueOf(finalBillAmount), "text"));
                        textParaRunIndexHashMap.put(invoiceDatePair, new Pair<>(invoiceDate, "table"));
                        textParaRunIndexHashMap.put(softexNumberPair, new Pair<>(prevF, "table"));
                        textParaRunIndexHashMap.put(nameOfBuyerPair, new Pair<>(currentD, "table"));

                        // Document updates
                        docxFileOperations.updateTextAtPosition(excelFilePath, currentWorkingDirectory, textParaRunIndexHashMap);

                        heirarchyIndex++;
                        invoiceDate = "";
                        currentWorkingDirectory = folderOperations.createFolder(String.valueOf(heirarchyIndex), pathBefore(currentWorkingDirectory)); // create folder with name = '1'
                        billAmount = cellValBillAmount;
                        chargesAmount = cellValChargesAmount;
                        finalBillAmount = cellValFinalBillAmount;
                    }
                } else {
                    billAmount = Math.round(billAmount * 100.0) / 100.0;
                    chargesAmount = Math.round(chargesAmount * 100.0) / 100.0;
                    finalBillAmount = Math.round(finalBillAmount * 100.0) / 100.0;

                    textParaRunIndexHashMap.put(billAmountPair, new Pair<>(String.valueOf(billAmount), "text"));
                    textParaRunIndexHashMap.put(chargesAmountPair, new Pair<>(String.valueOf(chargesAmount), "text"));
                    textParaRunIndexHashMap.put(finalBillAmountPair, new Pair<>(String.valueOf(finalBillAmount), "text"));
                    textParaRunIndexHashMap.put(invoiceDatePair, new Pair<>(invoiceDate, "table"));
                    textParaRunIndexHashMap.put(softexNumberPair, new Pair<>(prevF, "table"));
                    textParaRunIndexHashMap.put(nameOfBuyerPair, new Pair<>(prevD, "table"));

                    // Document updates
                    docxFileOperations.updateTextAtPosition(excelFilePath, currentWorkingDirectory, textParaRunIndexHashMap);

                    String path = folderOperations.createFolder(currentD, resultantFilePath);
                    heirarchyIndex = 1;
                    currentWorkingDirectory = folderOperations.createFolder(String.valueOf(heirarchyIndex), path); // create folder with name = '1'
                    invoiceDate = "";
                    billAmount = cellValBillAmount;
                    chargesAmount = cellValChargesAmount;
                    finalBillAmount = cellValFinalBillAmount;
                }
                // Invoice pdf search and save in current working directory
                File invoiceFile = folderOperations.searchForFile(GenApplication.invoiceDirectoriesPath, currentCellC.toString() + ".pdf");
                folderOperations.saveFileToOutputPath(invoiceFile, currentWorkingDirectory);

                prevD = currentD;
                prevF = currentF;
                prevB = currentB;
            }

        }

        System.out.println("\u001B[35m" + getClass().getName() + " : Operation Completed without exception !\u001B[0m" );
    }
}


