package com.chicmic.Util.DocumentOperations;


import com.chicmic.ExcelReadAndDataTransfer.ExcelPerformOperations;
import com.chicmic.engine.MainRunner;
import org.apache.commons.math3.util.Pair;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;


import java.io.*;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
public class DocxFileOperations {

    public Class<?> getCallingClass(int level) {
        StackTraceElement[] stackTraceElements = Thread.currentThread().getStackTrace();
        if (stackTraceElements.length >= (3 + level)) {
            StackTraceElement caller = stackTraceElements[2 + level];
            try {
                return Class.forName(caller.getClassName());
            } catch (ClassNotFoundException e) {
                e.printStackTrace();
            }
        }
        return null;
    }
    public void getParagraphAndRunIndices(String inputFilePath) throws IOException {

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
    public XWPFDocument insertTableInParagraph(XWPFDocument document, int paragraphIndex) throws IOException, InvalidFormatException {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        XWPFParagraph targetParagraph = paragraphs.get(paragraphIndex);

        FileInputStream fis = new FileInputStream(MainRunner.rootDirectory + "/" + MainRunner.FILE_NAME);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        int rows = sheet.getPhysicalNumberOfRows();
        int columns = 9;
        // Create a new table with gray-colored headings
        XWPFTable table = document.createTable(rows, columns);
        CTTbl ttbl = table.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr();
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
        CTBorder border = borders.addNewBottom();
        border.setColor("auto");
        border.setSz(BigInteger.valueOf(4));

        for (int row = 0; row < rows; row++) {
            XWPFTableRow tableRow = table.getRow(row);
            if (sheet.getRow(row) == null) {
                continue;
            }
            for (int col = 0; col < columns; col++) {
                XWPFTableCell cell = tableRow.getCell(col);
                XWPFParagraph cellParagraph = cell.getParagraphs().get(0);
                XWPFRun cellRun = cellParagraph.createRun();
                cellRun.setFontFamily("Arial");
                cellRun.setFontSize(7);

                if (row == 0) {
                    cellRun.setBold(true);
                    cell.setColor("000000"); // Black color
                    cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    CTTcPr cellPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
                    CTShd shd = cellPr.isSetShd() ? cellPr.getShd() : cellPr.addNewShd();
                    shd.setFill("BFBFBF"); // Grey color
//                    cellRun.setColor("808080"); // Gray color (use other color codes)
//                    cellRun.setBold(true);
                }

                cellParagraph.setSpacingAfter(0);
                cellParagraph.setSpacingBefore(0);

                String cellValue = sheet.getRow(row).getCell(col).toString();
                if (row > 0 && col >= 6) {
                    if(cellValue == null || cellValue.isEmpty()){
                        continue;
                    }
                    double cellDoubleValue = Double.parseDouble(cellValue);
                    cellDoubleValue = Math.round(cellDoubleValue * 100.0) / 100.0;
                    cellValue = String.valueOf(cellDoubleValue);
                }
                if (row > 0 && col == 0) cellValue = String.valueOf(row);
                cellRun.setText(cellValue);

            }
        }

        targetParagraph.getBody().insertNewTbl(targetParagraph.getCTP().newCursor()).getCTTbl().set(ttbl);

        fis.close();

        return document;
    }

    public void updateTextAtPosition(String inputFilePath, String outputFilePath, HashMap<Pair<Integer, Integer>, Pair<String, String>> textParaRunIndexMap) throws IOException {
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
                    run.setText(newText, 0);
                    if ((getCallingClass(1) == ExcelPerformOperations.class) && (paragraphIndex == 1 && runIndex == 2) || (paragraphIndex == 31 && runIndex == 1)) {
                        clearRunsInRange(document, paragraphIndex, runIndex + 1);
                    }
//                    else if((getCallingClass(1) == FourPointDeclaration.class) && (paragraphIndex == 3 && runIndex == 3)){
//                        clearRunsInRange(document, paragraphIndex, 4, 5);
//                    }

                }
            } else if (docType.startsWith("table")) {
                if (docType.equals("table_add")) {
                    System.out.println("Table cell update");
                    try {
                        document = insertTableInParagraph(document, paragraphIndex);
                    } catch (InvalidFormatException e) {
                        throw new RuntimeException(e);
                    }
                } else {
                    int tableIndex;
                    try {
                        tableIndex = Integer.parseInt(docType.substring("table".length()));
                    } catch (NumberFormatException e) {
                        throw new IllegalArgumentException("Invalid table index in docType: " + docType);
                    }
                    updateTableText(document, paragraphIndex, runIndex, tableIndex, newText);
                }
            }

        }
        System.out.println("\u001B[35m Document Updated Path " + outputFilePath + "\u001B[0m");
        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
        fileInputStream.close();
    }

    private void updateTableText(XWPFDocument document, int rowIndex, int colIndex, int tableIndex, String newText) {
        XWPFTable table = document.getTableArray(tableIndex);
        if (rowIndex < table.getNumberOfRows()) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (colIndex < row.getTableCells().size()) {
                XWPFTableCell cell = row.getCell(colIndex);
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                        paragraph.removeRun(i);
                    }
                }
                cell.setText(newText);
            }
        }
    }


    public void clearRunsInRange(XWPFDocument document, int paragraphIndex, int startRunIndex) {
        if (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size()) {
            XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);
            int totalRuns = paragraph.getRuns().size();
            startRunIndex = Math.max(0, startRunIndex);

            for (int i = startRunIndex; i < totalRuns; i++) {
                XWPFRun run = paragraph.getRuns().get(i);
                run.setText("", 0);
            }
        }
    }
    public void clearRunsInRange(XWPFDocument document, int paragraphIndex, int startRunIndex, int endRunIndex) {
        if (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size()) {
            XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);
            int totalRuns = paragraph.getRuns().size();

            startRunIndex = Math.max(0, startRunIndex);
            endRunIndex = Math.min(endRunIndex, totalRuns - 1);

            for (int i = startRunIndex; i <= endRunIndex; i++) {
                XWPFRun run = paragraph.getRuns().get(i);
                run.setText("", 0); // Clear the text of the run
            }
        }
    }
    public void deleteTableAtIndex(String inputFilePath, String outputFilePath, int tableIndex) throws IOException {

        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        if (tableIndex >= 0 && tableIndex < document.getTables().size()) {

            document.removeBodyElement(tableIndex + document.getParagraphs().size()); // Adjust for paragraphs
        } else {
            System.out.println("Table index is out of range.");
        }

        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
        fileInputStream.close();
        System.out.println("Table at index " + tableIndex + " deleted. Document updated at " + outputFilePath);
    }

    public void deleteTables(String inputFilePath) {
        try {
            FileInputStream fileInputStream = new FileInputStream(inputFilePath);
            XWPFDocument document = new XWPFDocument(fileInputStream);
            int numTables = document.getTables().size();
            for (int i = numTables - 1; i >= 0; i--) {
                XWPFTable table = document.getTables().get(i);
                document.removeBodyElement(document.getPosOfTable(table));
            }

            try (FileOutputStream fos = new FileOutputStream(inputFilePath + "modified_document.docx")) {
                document.write(fos);
            }

            System.out.println("All tables have been deleted from the document while preserving the text.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



}
