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

        System.out.println("hashmap = " + textParaRunIndexHashMap.size());

        FolderOperations folderOperations = new FolderOperations();
        DocxFileOperations docxFileOperations = new DocxFileOperations();

        String excelFilePath = rootDirectoryPath;
        System.out.println("\u0001B35m excelFilePath = " + excelFilePath);
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

        System.err.println("File Path exper = " + excelFile.getAbsolutePath());
        // use this to find the document paragraph index and run index by entering path of doc file
        docxFileOperations.getParagraphAndRunIndices(excelFilePath);

        String prevD = "";
        String prevF = "";
        int heirarchyIndex = 0;
        String currentWorkingDirectory = resultantFilePath;
        for (int i = 0; i < rows.size(); i++) {
            System.out.println("\u001B[37m Currenting working directory " + currentWorkingDirectory + "\u001B[0m");
            Row sortedRow = rows.get(i);
            Cell currentCellA = sortedRow.getCell(0);
            Cell currentCellB = sortedRow.getCell(1);
            Cell currentCellC = sortedRow.getCell(2);
            Cell currentCellD = sortedRow.getCell(indexOfRecipientColumnD);
            Cell currentCellF = sortedRow.getCell(indexOfRecipientColumnD + 2);
            System.out.println("cell value = "+ currentCellA + " " + currentCellB + " " + currentCellC + " " + currentCellD);

            if (currentCellD != null) {
                String currentD = currentCellD.toString();
                String currentF = currentCellF.toString();
                String currentB = currentCellB.toString();
                String invoiceDate = currentB;
                double cellValBillAmount = Double.parseDouble(sortedRow.getCell(indexOfRecipientColumnD + 3).toString());
                double cellValChargesAmount = Double.parseDouble(sortedRow.getCell(indexOfRecipientColumnD + 4).toString());
                double cellValFinalBillAmount = Double.parseDouble(sortedRow.getCell(indexOfRecipientColumnD + 5).toString());

                DecimalFormat decimalFormat = new DecimalFormat("#.##");

                // Round off the double values to two decimal places
                cellValBillAmount = Double.parseDouble(decimalFormat.format(cellValBillAmount));
                cellValChargesAmount = Double.parseDouble(decimalFormat.format(cellValChargesAmount));
                cellValFinalBillAmount = Double.parseDouble(decimalFormat.format(cellValFinalBillAmount));
                System.out.println("\u001B[33m val = prevD " + prevD + " currentD= " + currentD + " Ano. : " + currentCellA.toString() +  "\u001B[0m");
                if (prevD.equals(currentD)) {
                    System.out.println("index = " + i);
                    if (currentF.equals(prevF)) {
//                            System.out.println("\u001B[34m in prevF if "  + currentF + "\u001B[0m");

                        billAmount += cellValBillAmount;
                        chargesAmount += cellValChargesAmount;
                        finalBillAmount += cellValFinalBillAmount;
                    } else {

//                        textParaRunIndexHashMap.put(invoiceDate, new Pair<>(10, 3));
                        System.out.println("\u001B[34m size= " + textParaRunIndexHashMap.size());
                        for (Map.Entry<Pair<Integer, Integer>, Pair<String, String>> entry : textParaRunIndexHashMap.entrySet()) {
                            Pair<Integer, Integer> key = entry.getKey();
                            String value = String.valueOf(entry.getValue());
                            System.out.println("Key: " + key + ", Value: " + value);
                        }
                        textParaRunIndexHashMap.put(billAmountPair, new Pair<>(String.valueOf(billAmount), "text"));
                        textParaRunIndexHashMap.put(chargesAmountPair, new Pair<>(String.valueOf(chargesAmount), "text"));
                        textParaRunIndexHashMap.put(finalBillAmountPair, new Pair<>(String.valueOf(finalBillAmount), "text"));
                        textParaRunIndexHashMap.put(invoiceDatePair, new Pair<>(invoiceDate, "table"));
                        textParaRunIndexHashMap.put(softexNumberPair, new Pair<>(currentF, "table"));
                        textParaRunIndexHashMap.put(nameOfBuyerPair, new Pair<>(currentD, "table"));



                        System.out.println("\u001B[0m inside else excelfilePath == " + excelFilePath);
                        docxFileOperations.updateTextAtPosition(excelFilePath, currentWorkingDirectory, textParaRunIndexHashMap);
                        heirarchyIndex++;
                        currentWorkingDirectory = folderOperations.createFolder(String.valueOf(heirarchyIndex), pathBefore(currentWorkingDirectory)); // create folder with name = '1'
                    }
//                        System.out.println("currentD = " + currentD+  " currentF = " + currentF.toString());
                } else {
                    System.out.println("index = " + i);
                    textParaRunIndexHashMap.put(billAmountPair, new Pair<>(String.valueOf(billAmount), "text"));
                    textParaRunIndexHashMap.put(chargesAmountPair, new Pair<>(String.valueOf(chargesAmount), "text"));
                    textParaRunIndexHashMap.put(finalBillAmountPair, new Pair<>(String.valueOf(finalBillAmount), "text"));
                    textParaRunIndexHashMap.put(invoiceDatePair, new Pair<>(invoiceDate, "table"));
                    textParaRunIndexHashMap.put(softexNumberPair, new Pair<>(currentF, "table"));
                    textParaRunIndexHashMap.put(nameOfBuyerPair, new Pair<>(currentD, "table"));

                    docxFileOperations.updateTextAtPosition(excelFilePath, currentWorkingDirectory, textParaRunIndexHashMap);

//                        System.out.println("Else currentD = " + currentD+  " currentF = " + currentF.toString());
                    String path = folderOperations.createFolder(currentD, resultantFilePath);
                    heirarchyIndex = 1;
                    currentWorkingDirectory = folderOperations.createFolder(String.valueOf(heirarchyIndex), path); // create folder with name = '1'


                    billAmount = cellValBillAmount;
                    chargesAmount = cellValChargesAmount;
                    finalBillAmount = cellValFinalBillAmount;
                }
                File invoiceFile = folderOperations.searchForFile(rootDirectoryPath, currentCellC.toString() + ".pdf");
                folderOperations.saveFileToOutputPath(invoiceFile, currentWorkingDirectory);

                prevD = currentD;
                prevF = currentF;
            }

        }

        System.out.println("Excel sheet sorted, column G updated, and new file generated based on column D.");
    }
}


