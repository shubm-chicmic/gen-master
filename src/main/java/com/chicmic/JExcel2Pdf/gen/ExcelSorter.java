package com.chicmic.JExcel2Pdf.gen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelSorter {


    public File excelManager(File file) {

        file = excelSortByColumn(file, 3);
        return file;
    }

    public File excelSortByColumn(File file, int columnIndex) {
        File sortedFile = null;
        try {
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0); // Assuming it's the first sheet

            // Extract rows to a list
            List<Row> rows = new ArrayList<>();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                rows.add(row);
            }

            Comparator<Row> comparator = (r1, r2) -> {
                Cell cell1 = r1.getCell(columnIndex);
                Cell cell2 = r2.getCell(columnIndex);
                return cell1.toString().compareTo(cell2.toString());
            };
            rows.sort(comparator);

            // Iterate through the sorted rows and sort rows with equal values in column 3
            int start = 0;
            while (start < rows.size()) {
                int end = start + 1;
                while (end < rows.size() && rows.get(end).getCell(columnIndex).toString().equals(rows.get(start).getCell(columnIndex).toString())) {
                    end++;
                }

                // Sort rows within the equal value range in column 3
                List<Row> subList = rows.subList(start, end);
                Comparator<Row> subComparator = (r1, r2) -> {
                    Cell cell1 = r1.getCell(5); // Sorting by column 5 within the equal value range of column 3
                    Cell cell2 = r2.getCell(5);
                    return cell1.toString().compareTo(cell2.toString());
                };
                subList.sort(subComparator);

                start = end;
            }

            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("Sorted Data");

            int rowIndex = 0;
            for (Row sortedRow : rows) {
                Row newRow = newSheet.createRow(rowIndex++);

                for (int j = 0; j < sortedRow.getLastCellNum(); j++) {
                    Cell cell = newRow.createCell(j);
                    Cell originalCell = sortedRow.getCell(j);

                    if (originalCell != null) {
                        cell.setCellValue(originalCell.toString());
                    }
                }
            }

            // Generate a unique file name for the sorted Excel file
            String originalFileName = file.getName();
            String sortedFileName = "sorted_" + originalFileName;
            sortedFile = new File(sortedFileName);

            // Write the new workbook to the sortedFile
            FileOutputStream fos = new FileOutputStream(sortedFile);
            newWorkbook.write(fos);
            fos.close();

            fis.close();

            System.out.println(getClass().getName() +" : Excel sheet sorted based on column " + columnIndex + " and a new file generated: " + sortedFileName);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return sortedFile;
    }







}
