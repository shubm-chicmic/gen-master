package com.chicmic.ExcelReadAndDataTransfer;

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
            Sheet sheet = workbook.getSheetAt(0);

            // Extract rows to a list, but skip null rows and empty cells
            List<Row> rows = new ArrayList<>();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    boolean hasCellValue = false;  // Flag to check if any cell has a value
                    for (int cellIndex = 0; cellIndex < 9; cellIndex++) {
                        Cell cell = row.getCell(cellIndex);
                        if (cell != null) {
                            if (cell.getCellType() == CellType.STRING) {
                                String cellValue = cell.getStringCellValue();
                                if (cellValue != null && !cellValue.isEmpty()) {
                                    hasCellValue = true;
                                    break;
                                }
                            } else if (cell.getCellType() != CellType.BLANK) {
                                hasCellValue = true;
                                break;
                            }
                        }
                    }

                    if (hasCellValue) {
                        rows.add(row);
                    } else {
                        System.out.println("Empty row at index: " + rowIndex);
                    }
                } else {
                    System.out.println("Null row at index: " + rowIndex);
                    break;
                }
            }
            Comparator<Row> comparator = (r1, r2) -> {
                if (r1 == null && r2 == null) {
                    return 0;
                } else if (r1 == null) {
                    return 1;
                } else if (r2 == null) {
                    return -1;
                } else {
                    Cell cell1 = r1.getCell(columnIndex);
                    Cell cell2 = r2.getCell(columnIndex);
                    if (cell1 == null) {
                        System.out.println("Cell at row " + r1.getRowNum() + " and column " + columnIndex + " is null");
                    }
                    if (cell2 == null) {
                        System.out.println("Cell at row " + r2.getRowNum() + " and column " + columnIndex + " is null");
                    }
                    if (cell1 != null && cell2 != null) {
                        return cell1.toString().compareTo(cell2.toString());
                    } else {
                        return 0; // Handle null cells as equal
                    }
                }
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
                Row newRow = newSheet.createRow(rowIndex);
                rowIndex++;

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
