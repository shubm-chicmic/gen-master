package com.chicmic.engine;

import com.chicmic.pdfOperations.DocUpdateWithPdfData;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class PdfReadAndDocUpdate {
    public static void main(String[] args) {
        DocUpdateWithPdfData docUpdateWithPdfData = new DocUpdateWithPdfData();
//        docUpdateWithPdfData.documentUpdate();
        readPDF("sample1.pdf");
    }
    public static void readPDF(String filePath) {
        try (PDDocument document = PDDocument.load(new File(filePath))) {
            PDFTextStripper stripper = new PDFTextStripper() {
                @Override
                protected void startDocument(PDDocument document) throws IOException {
                    super.startDocument(document);
                    System.out.println("PDF Document Information:");
                    System.out.println("Number of Pages: " + document.getNumberOfPages());
                    System.out.println("---------------------------");
                }

                @Override
                protected void startPage(PDPage page) throws IOException {
                    super.startPage(page);
                    int pageNumber = getCurrentPageNo();
                    System.out.println("Page " + pageNumber + ":\n---------------------------");

                }

                @Override
                protected void writeString(String text, List<TextPosition> textPositions) throws IOException {
                    super.writeString(text, textPositions);
                    System.out.println(text); // Print text content of each page
//                    String[] words = text.split("\\s+"); // Split text into words
//
//                    for (String word : words) {
//                        // Find positions for each word
//                        for (TextPosition textPosition : textPositions) {
//                            // Check if the text position contains the current word
//                            if (textPosition.toString().contains(word)) {
//                                System.out.println("Word: " + word);
//                                System.out.println("Position: " + textPosition.getXDirAdj() + ", " + textPosition.getYDirAdj());
//                                break; // Break the loop after finding the word's position
//                            }
//                        }
//                    }
                }

                @Override
                protected void endPage(PDPage page) throws IOException {
                    super.endPage(page);
                    System.out.println("\n---------------------------\n"); // Separator between pages
                }

                @Override
                protected void endDocument(PDDocument document) throws IOException {
                    super.endDocument(document);
                    System.out.println("End of Document");
                }
            };

            // Set configurations for text extraction (optional)
            stripper.setSortByPosition(true); // Process text based on its position

            // Extract text from the PDF
            stripper.getText(document);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Process the extracted table data
    private static void processTable(List<List<String>> table) {
        for (List<String> row : table) {
            for (String cell : row) {
                System.out.print(cell + "\t");
            }
            System.out.println();
        }
    }
}
