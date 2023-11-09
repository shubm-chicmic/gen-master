package com.chicmic.JExcel2Pdf.gen;


import org.apache.commons.math3.util.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;

import java.io.*;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;

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

        }
    }

    public void addTableAfterParagraphIndex(String inputFilePath, String outputFilePath, int paragraphIndex, File excelFile) throws IOException {
        FileInputStream fis = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fis);
        fis.close();

        // Create a new paragraph after the specified index
        XWPFParagraph paragraph = document.createParagraph();
        document.setParagraph(paragraph, paragraphIndex + 1);

        // Read data from Excel file and create a table in the new paragraph
        Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFile));
        Sheet sheet = workbook.getSheetAt(0);

        XWPFTable table = document.createTable();
        int rowCount = sheet.getPhysicalNumberOfRows();

        // Iterate through Excel rows and populate the table
        for (int i = 0; i < rowCount; i++) {
            Row excelRow = sheet.getRow(i);
            XWPFTableRow tableRow = table.createRow();
            if (excelRow == null) {
                break;
            }
            int columnCount = excelRow.getPhysicalNumberOfCells();
            for (int j = 0; j < columnCount; j++) {
                Cell excelCell = excelRow.getCell(j);
                XWPFTableCell tableCell = tableRow.createCell();
                if (excelCell != null) {
                    tableCell.setText(excelCell.toString());
                } else {
                    tableCell.setText("");
                }
            }
        }

        FileOutputStream fos = new FileOutputStream(outputFilePath);
        document.write(fos);
        fos.close();

        System.out.println("Table added after paragraph index " + paragraphIndex + ". Document updated at " + outputFilePath);
    }

    public void insertTableAtIndex(String inputDocxFilePath, String outputDocxFilePath, File excelFile, int tableIndex) throws IOException {

        FileInputStream fileInputStream = new FileInputStream(inputDocxFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        if (tableIndex < 0 || tableIndex > document.getTables().size()) {
            System.out.println("Table index is out of range.");
            return;
        }

        if (excelFile != null && excelFile.exists()) {
            FileInputStream fis = new FileInputStream(excelFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0); // Assuming you want the first sheet

            XWPFTable table = document.getTables().get(tableIndex);
            List<XWPFTableRow> tableRows = table.getRows();

            for (int rowIndex = 0; rowIndex < sheet.getLastRowNum() + 1; rowIndex++) {
                XWPFTableRow tableRow;
                if (rowIndex < tableRows.size()) {
                    tableRow = tableRows.get(rowIndex);
                } else {
                    tableRow = table.createRow();
                }

                XSSFRow excelRow = (XSSFRow) sheet.getRow(rowIndex);
                if (excelRow != null) {
                    for (int cellIndex = 0; cellIndex < excelRow.getLastCellNum(); cellIndex++) {
                        XWPFTableCell tableCell = tableRow.getCell(cellIndex);
                        if (tableCell == null) {
                            tableCell = tableRow.createCell();
                        }
                        XSSFCell excelCell = excelRow.getCell(cellIndex);
                        if (excelCell != null) {
                            String cellValue = excelCell.toString();
                            if (rowIndex > 0 && cellIndex >= 6 && cellIndex <= 8) {
                                try {
                                    double cellDoubleValue = Double.parseDouble(cellValue);
                                    cellValue = String.format("%.2f", cellDoubleValue);
                                } catch (NumberFormatException e) {

                                }
                            }
                            tableCell.setText("");
                        }
                    }
                }
            }

            System.out.println("New table (copy of Excel file) inserted at index " + tableIndex + ".");
        } else {
            System.out.println("Excel file is not provided or does not exist.");
        }

        FileOutputStream fileOutputStream = new FileOutputStream(outputDocxFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
        fileInputStream.close();
    }

    public XWPFDocument insertTableInParagraph(XWPFDocument document, int paragraphIndex, int runIndex) throws IOException, InvalidFormatException {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        XWPFParagraph targetParagraph = paragraphs.get(paragraphIndex);

        // Load the data from the Excel file
        FileInputStream fis = new FileInputStream(GenApplication.rootDirectory + "/" + GenApplication.FILE_NAME);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        int rows = 50;
        int columns = 9;
        // Create a new table with gray-colored headings
        XWPFTable table = document.createTable(rows + 1, columns);
        CTTbl ttbl = table.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr();
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
        CTBorder border = borders.addNewBottom();
        border.setColor("auto");
        border.setSz(BigInteger.valueOf(4));


        // Populate the table with data from the Excel sheet
        for (int row = 0; row < rows; row++) {
            if (sheet.getRow(row) == null) {
                continue;
            }
            for (int col = 0; col < columns; col++) {
                XWPFTableCell cell = table.getRow(row + 1).getCell(col);
                XWPFParagraph cellParagraph = cell.getParagraphs().get(0);
                XWPFRun cellRun = cellParagraph.createRun();
                if (row == 0) {
                    cellRun.setColor("808080"); // Gray color (you can use other color codes)
                    cellRun.setBold(true); // Make the heading bold
                }
                cellRun.setText(sheet.getRow(row).getCell(col).toString());
            }
        }

        targetParagraph.getBody().insertNewTbl(targetParagraph.getCTP().newCursor()).getCTTbl().set(ttbl);

        fis.close();

        return document;
    }

    public void updateTextAtPosition(String inputFilePath, String outputFilePath, HashMap<Pair<Integer, Integer>, Pair<String, String>> textParaRunIndexMap) throws IOException {

//        System.out.println("\u001B[33m document path = " + inputFilePath + "\u001B[0m");
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
            } else if (docType.equals("table")) {
                // Handle table cell updates
                updateTableText(document, paragraphIndex, runIndex, newText);
            } else if (docType.equals("table_add")) {
                try {
                    System.out.println("Table cell update");
                    document = insertTableInParagraph(document, paragraphIndex, runIndex);
                } catch (InvalidFormatException e) {
                    throw new RuntimeException(e);
                }
            }

        }
        System.out.println("\u001B[35m Document Updated Path " + outputFilePath + "\u001B[0m");
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

            startRunIndex = Math.max(0, startRunIndex);


            for (int i = startRunIndex; i < totalRuns; i++) {
                XWPFRun run = paragraph.getRuns().get(i);
                run.setText("", 0); // Clear the text of the run
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

}
