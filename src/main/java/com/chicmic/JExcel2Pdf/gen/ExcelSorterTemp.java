package com.chicmic.JExcel2Pdf.gen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class ExcelSorterTemp {
    public static void excelReadAndSort2(File file) {
        try {
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0); // Assuming it's the first sheet

            // Create a custom comparator for sorting by column D (0-based index)
            int columnIndexToSort = 3; // Column D is index 3 (0-based index)

            Comparator<Row> comparator = (r1, r2) -> {
                Cell cell1 = r1.getCell(columnIndexToSort);
                Cell cell2 = r2.getCell(columnIndexToSort);
                return cell1.toString().compareTo(cell2.toString());
            };

            // Convert the sheet's rows to a list for sorting, skipping the first row
            List<Row> rows = new ArrayList<>();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                rows.add(row);
            }

            // Sort the rows using the custom comparator
            rows.sort(comparator);

            int rowIndex = 0;
            double sum = 0.0; // Initialize the sum
            String currentDValue = "";

            for (int i = 0; i < rows.size(); i++) {
                Row sortedRow = rows.get(i);

                // Check if the values in column D are the same as the next row and update the sum
                Cell currentCellD = sortedRow.getCell(columnIndexToSort);
                if (currentCellD != null) {
                    String currentD = currentCellD.toString();

                    if (currentD.equals(currentDValue)) {
                        // Get the value in column G and add it to the sum
                        Cell currentCellG = sortedRow.getCell(6); // Assuming G is column 7 (0-based index)
                        if (currentCellG != null && currentCellG.getCellType() == CellType.NUMERIC) {
                            double currentCellValueG = currentCellG.getNumericCellValue();
                            sum += currentCellValueG;
                        }
                    } else {
                        // Display the sum for the previous set of rows with the same D value
                        if (!currentDValue.isEmpty()) {
                            System.out.println("For D = " + currentDValue + ", Sum of G = " + sum);
                        }

                        // Set the new D value and reset the sum
                        currentDValue = currentD;
                        sum = 0.0;
                    }
                }
            }

            // Display the sum for the last set of rows with the same D value
            if (!currentDValue.isEmpty()) {
                System.out.println("For D = " + currentDValue + ", Sum of G = " + sum);
            }

            // Close the input file
            fis.close();

            System.out.println("Excel sheet sorted, and sums displayed based on column D.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }



}
