package com.chicmic.JExcel2Pdf.gen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class ExcelRead {
    public static void excelReadAndSort(File excelFile){
        try {

            FileInputStream fis = new FileInputStream(excelFile);
            Workbook workbook = new XSSFWorkbook(fis);
            fis.close();

            Sheet sheet = workbook.getSheetAt(0);

            // Read and store data from columns D and M
            List<RowData> rowDataList = new ArrayList<>();
            for (Row row : sheet) {
                Cell cellD = row.getCell(3); // Column D is 0-based index 3
                Cell cellF = row.getCell(5); // Column F is 0-based index 5

                if (cellD != null && cellF != null) {
                    System.out.println("REadomg " + cellD.getStringCellValue());
                    rowDataList.add(new RowData(
                            getCellStringValue(cellD),
                            getCellStringValue(cellF)
                    ));
                }
            }
            System.out.println("Before Sorting:");
            for (RowData rowData : rowDataList) {
                System.out.println("D: " + rowData.columnD + ", F: " + rowData.columnF);
            }

            // Sort the data based on column D
            Collections.sort(rowDataList, (row1, row2) -> row1.columnD.compareTo(row2.columnD));
            System.out.println("\nAfter Sorting:");
            for (RowData rowData : rowDataList) {
                System.out.println("D: " + rowData.columnD + ", F: " + rowData.columnF);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    static class RowData {
        String columnD;
        String columnF;

        RowData(String columnD, String columnF) {
            this.columnD = columnD;
            this.columnF = columnF;
        }
    }

    static String getCellStringValue(Cell cell) {

        if (CellType.STRING.equals(cell.getCellType())) {
            return cell.getStringCellValue();
        } else if (CellType.NUMERIC.equals(cell.getCellType())) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return "";
        }
    }
}
