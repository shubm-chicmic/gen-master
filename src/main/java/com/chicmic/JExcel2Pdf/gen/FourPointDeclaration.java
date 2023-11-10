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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class FourPointDeclaration {

    Pair<Integer, Integer> minDateAndMaxDateIndex = new Pair<>(3, 3);
    Pair<Integer, Integer> billAmountIndex1 = new Pair<>(3, 6);
    Pair<Integer, Integer> billAmountIndex2 = new Pair<>(5, 3);
    Pair<Integer, Integer> FINWNumberIndex = new Pair<>(5, 6);
    Pair<Integer, Integer> finalBillAmountIndex = new Pair<>(7, 1);
    Pair<Integer, Integer> tableIndex = new Pair<>(6, 0);

    HashMap<Pair<Integer, Integer>, Pair<String, String>> documentIndexAndTextMap = new HashMap<>();
    private final DocxFileOperations docxFileOperations = new DocxFileOperations();
    private final FolderOperations folderOperations = new FolderOperations();
    String templateDocument = GenApplication.fourPointDeclarationDocumentPath;

    public void generateDocument(File excelFile) throws IOException {

//        docxFileOperations.getParagraphAndRunIndices(templateDocument);

        FileInputStream fis = new FileInputStream(excelFile);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);


        List<Row> rows = new ArrayList<>();
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            rows.add(row);
        }

        String minDate = "";
        String maxDate = "";
        double billAmountTotal = 0.0;
        double finalBillAmountTotal = 0.0;
        String finalDocumentPath;
        for (Row excelRows : rows) {
            if (excelRows != null) {
                Cell cellB = excelRows.getCell(1); // Cell For Date
                Cell cellG = excelRows.getCell(6); // Cell For BillAmount
                Cell cellI = excelRows.getCell(8); // Cell For FinalBillAmount
                if (cellB == null || cellG == null || cellI == null) {
                    continue;
                }
                minDate = DateOperations.findMinimumDate(cellB.toString(), minDate);
                maxDate = DateOperations.findMaximumDate(cellB.toString(), maxDate);
                double cellBillAmount = Double.parseDouble(cellG.toString());
                double cellFinalBillAmount = Double.parseDouble(cellI.toString());
                cellBillAmount = Math.round(cellBillAmount * 100.0) / 100.0;
                cellFinalBillAmount = Math.round(cellFinalBillAmount * 100.0) / 100.0;

                billAmountTotal += cellBillAmount;
                finalBillAmountTotal += cellFinalBillAmount;

            }
        }


        billAmountTotal = Math.round(billAmountTotal * 100.0) / 100.0;
        finalBillAmountTotal = Math.round(finalBillAmountTotal * 100.0) / 100.0;

        documentIndexAndTextMap.put(billAmountIndex1, new Pair<>(String.valueOf(billAmountTotal), "text"));
        documentIndexAndTextMap.put(billAmountIndex2, new Pair<>(String.valueOf(billAmountTotal), "text"));
        documentIndexAndTextMap.put(FINWNumberIndex, new Pair<>("", "text"));
        documentIndexAndTextMap.put(finalBillAmountIndex, new Pair<>("USD " + (finalBillAmountTotal), "text"));
        documentIndexAndTextMap.put(minDateAndMaxDateIndex, new Pair<>(minDate + " till " + maxDate, "text"));
        documentIndexAndTextMap.put(tableIndex, new Pair<>("", "table_add"));

        finalDocumentPath = folderOperations.createFolder("FourPointDeclaration", GenApplication.rootDirectory);
        finalDocumentPath += "/FourPointDeclaration.docx";
        docxFileOperations.updateTextAtPosition(templateDocument, finalDocumentPath, documentIndexAndTextMap);
        docxFileOperations.deleteTableAtIndex(finalDocumentPath, finalDocumentPath, 1);
    }
}
