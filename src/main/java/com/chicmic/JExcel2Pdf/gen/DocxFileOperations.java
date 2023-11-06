package com.chicmic.JExcel2Pdf.gen;


import org.apache.commons.math3.util.Pair;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

public class DocxFileOperations {
    String updatedDocumentName = "L1 Request letter for Submission of Export doc.docx";

    public  void updateTextAtPosition(String inputFilePath, String outputFilePath, int paragraphIndex, int runIndex, String newText) throws IOException {
        outputFilePath += "/" + updatedDocumentName;
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        // Check if the specified paragraph and run indices are within valid ranges
        if (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size()) {
            XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);
            if (runIndex >= 0 && runIndex < paragraph.getRuns().size()) {
                XWPFRun run = paragraph.getRuns().get(runIndex);
                run.setText(newText, 0); // Replace the existing text with the new text
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
    }
    public void updateTextAtPosition(String inputFilePath, String outputFilePath, HashMap<Pair<Integer, Integer>, Pair<String, String>> textParaRunIndexMap) throws IOException {
        outputFilePath += "/" + updatedDocumentName;
        inputFilePath += "/" + updatedDocumentName;
        System.out.println("\u001B[33m document path = " + inputFilePath + "\u001B[0m");
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        for (Pair<Integer, Integer> paraRunIndices : textParaRunIndexMap.keySet()) {
            int paragraphIndex = paraRunIndices.getFirst();
            int runIndex = paraRunIndices.getSecond();
            String newText = textParaRunIndexMap.get(paraRunIndices).getFirst();
            String docType = textParaRunIndexMap.get(paraRunIndices).getSecond();

            if (docType.equals("text") && (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size())) {
                XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);

                if (runIndex >= 0 && runIndex < paragraph.getRuns().size()) {
                    XWPFRun run = paragraph.getRuns().get(runIndex);
                    run.setText(newText, 0); // Replace the existing text with the new text
                    if((paragraphIndex == 1 && runIndex == 2) || (paragraphIndex == 31 && runIndex == 1)){
                        clearRunsInRange(document, paragraphIndex, runIndex + 1);
                    }
                }
            } else if(docType.equals("table")){
                System.out.println("\u001B[31m New Text " + newText + " " + paragraphIndex + " " + runIndex);
                // Handle table cell updates
                updateTableText(document, paragraphIndex, runIndex, newText);
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
        fileInputStream.close();
    }

    private void updateTableText(XWPFDocument document, int rowIndex, int colIndex, String newText) {
        for (XWPFTable table : document.getTables()) {
            if (rowIndex < table.getNumberOfRows()) {
                XWPFTableRow row = table.getRow(rowIndex);
                if (colIndex < row.getTableCells().size()) {
                    XWPFTableCell cell = row.getCell(colIndex);

                    // Clear the existing content of the cell
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                            paragraph.removeRun(i);
                        }
                    }

                    // Set the new text in the cell
                    cell.setText(newText);
                }
            }
        }
    }

    public void clearRunsInRange(XWPFDocument document, int paragraphIndex, int startRunIndex) {
        if (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size()) {
            XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);
            int totalRuns = paragraph.getRuns().size();

            // Ensure the startRunIndex and endRunIndex are within the valid range
            startRunIndex = Math.max(0, startRunIndex);


            // Clear the specified range of runs
            for (int i = startRunIndex; i < totalRuns; i++) {
                XWPFRun run = paragraph.getRuns().get(i);
                run.setText("", 0); // Clear the text of the run
            }
        }
    }

    public  void updateTextInRange(String inputFilePath, String outputFilePath, int paragraphIndex, int runStartIndex, int runEndIndex, String newText) throws IOException {
        outputFilePath += "/" + updatedDocumentName;

        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        // Check if the specified paragraph index is within a valid range
        if (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size()) {
            XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);

            // Ensure the run indices are within the valid range
            int maxRunIndex = paragraph.getRuns().size() - 1;
            if (runStartIndex >= 0 && runStartIndex <= maxRunIndex && runEndIndex >= runStartIndex && runEndIndex <= maxRunIndex) {
                // Create a new run with the updated text
                XWPFRun updatedRun = paragraph.insertNewRun(runStartIndex);
                updatedRun.setText(newText);

                // Remove the runs in the specified range
                for (int i = runStartIndex + 1; i <= runEndIndex; i++) {
                    paragraph.removeRun(runStartIndex + 1);
                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
    }
    public void getParagraphAndRunIndices(String inputFilePath) throws IOException {
        inputFilePath += "/" + updatedDocumentName;
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        int tableIndex = 0;
        for (XWPFTable table : document.getTables()) {
            System.out.println("\u001B[36mTable Index: " + tableIndex + "\u001B[0m");
            for (int rowIndex = 0; rowIndex < table.getRows().size(); rowIndex++) {
                XWPFTableRow row = table.getRow(rowIndex);
                System.out.println("\u001B[37mRow Index: " + rowIndex + "\u001B[0m");

                for (int colIndex = 0; colIndex < row.getTableCells().size(); colIndex++) {
                    XWPFTableCell cell = row.getCell(colIndex);
                    System.out.println("\u001B[38mColumn Index: " + colIndex + "\u001B[0m");
                    System.out.println("Cell Text: " + cell.getText() + "\u001B[0m"); // Print cell text
                }
            }
            tableIndex++;
        }

        int paragraphIndex = 0;
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            System.out.println("\u001B[34mParagraph Index: " + paragraphIndex + "\u001B[0m");

            int runIndex = 0;
            for (XWPFRun run : paragraph.getRuns()) {
                System.out.println("\u001B[35mRun Index: " + runIndex);
                System.out.println("Run Text: " + run.getText(0) + "\u001B[0m"); // Print run text
                runIndex++;
            }
            paragraphIndex++;
        }
    }
    public  void getParagraphAndRunIndices1(String inputFilePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.size(); paragraphIndex++) {
            XWPFParagraph paragraph = paragraphs.get(paragraphIndex);
            System.out.println("\u001B[34m Paragraph Index: " + paragraphIndex + "\u001B[0m");

            List<XWPFRun> runs = paragraph.getRuns();
            for (int runIndex = 0; runIndex < runs.size(); runIndex++) {
                XWPFRun run = runs.get(runIndex);
                System.out.println("\u001B[35m"+ "Run Index: " + runIndex);
                System.out.println("Run Text: " + run.getText(0) + "\u001B[0m"); // Print the text of the run
            }
        }
    }
}
