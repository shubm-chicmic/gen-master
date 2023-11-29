package com.chicmic.pdfOperations;

import com.chicmic.Util.DateOperations;
import com.chicmic.Util.DocumentOperations.DocxFileOperations;
import com.chicmic.Util.FolderOperations.FolderOperations;
import com.chicmic.engine.MainRunner;
import org.apache.commons.math3.util.Pair;

import java.io.IOException;
import java.util.HashMap;

public class DocUpdateWithPdfData {
    String documentName = MainRunner.exportRegularisationDocumentName;
    HashMap<Pair<Integer, Integer>, Pair<String, String>> documentIndexAndTextMap = new HashMap<>();
    private final DocxFileOperations docxFileOperations = new DocxFileOperations();
    private final FolderOperations folderOperations = new FolderOperations();
    private boolean isDoc = false;
    private final String currentDate = DateOperations.getTodaysDate();
    private String billAmount = "";
    private String softexNumber = "";
    private String FIRCNumber = "";
    private String accountNumber = "";
    private String importerExporterCode = "";
    private String invoiceNumber = "";

    public void documentUpdate() {
        try {
            if (documentName.endsWith(".doc")) {
                documentName = docxFileOperations.convertDocToDocx(documentName);
                isDoc = true;
            }
            docxFileOperations.getParagraphAndRunIndices(documentName);



            //Delete Temp File
            if (isDoc) {
                folderOperations.deleteFile(documentName);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    public void updateDocument() throws IOException {
        // Pair<Integer, Integer> currentDateIndex = new Pair<>(9, 7);
        documentIndexAndTextMap.put(new Pair<>(9, 7), new Pair<>(currentDate, "text"));
        // Pair<Integer, Integer> billAmountIndex = new Pair<>(19, 0);
        documentIndexAndTextMap.put(new Pair<>(19, 0), new Pair<>("Bill Amount USD" +billAmount, "text"));
        // Pair<Integer, Integer> softexNumberIndex = new Pair<>(21, 0);
        documentIndexAndTextMap.put(new Pair<>(21, 0), new Pair<>("GR / Shipping Bill / Softex Form No. " + softexNumber, "text"));
        // Pair<Integer, Integer> FIRCNumberIndex = new Pair<>(21, 2);
        documentIndexAndTextMap.put(new Pair<>(21, 2), new Pair<>("dated 19/10/2023 FIRC # " + FIRCNumber, "text"));
        // Pair<Integer, Integer> accountNumberIndex = new Pair<>(25, 0);
        documentIndexAndTextMap.put(new Pair<>(25, 0), new Pair<>("Debit all charges for processing of above-mentioned documents from account no " + accountNumber, "text"));
        // Pair<Integer, Integer> importerExporterCodeIndex = new Pair<>(27, 0);
        documentIndexAndTextMap.put(new Pair<>(27, 0), new Pair<>("We are eligible to export the above mentioned goods/services" +
                " under the current Foreign Trade policy in place. And our Importer Exporter Code is:    " +importerExporterCode, "text"));

        // TABLE UPDATE
        // Pair<Integer, Integer> invoiceNumberTableIndex1 = new Pair<>(1, 2);
        documentIndexAndTextMap.put(new Pair<>(1, 2), new Pair<>(invoiceNumber, "table1"));
        // Pair<Integer, Integer> FIRCNumberTableIndex1 = new Pair<>(1, 7);
        documentIndexAndTextMap.put(new Pair<>(1, 7), new Pair<>(FIRCNumber, "table1"));
        // Pair<Integer, Integer> softexNumberIndex1 = new Pair<>(1, 8);
        documentIndexAndTextMap.put(new Pair<>(1, 8), new Pair<>(softexNumber, "table1"));
        // Pair<Integer, Integer> invoiceNumberTableIndex2 = new Pair<>(2, 2);
        documentIndexAndTextMap.put(new Pair<>(2, 2), new Pair<>(invoiceNumber, "table1"));
        // Pair<Integer, Integer> FIRCNumberTableIndex2 = new Pair<>(2, 7);
        documentIndexAndTextMap.put(new Pair<>(2, 7), new Pair<>(FIRCNumber, "table1"));

        docxFileOperations.updateTextAtPosition(documentName, documentName + "update.docx", documentIndexAndTextMap);

    }
}
