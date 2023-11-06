package com.chicmic.JExcel2Pdf.gen;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

public class FolderOperations {
    public String createFolder(String folderName, String path) {
        File folder = new File(path, folderName);

        if (!folder.exists()) {
            boolean created = folder.mkdirs();
            if (created) {
                System.out.println("Folder created: " + folder.getAbsolutePath());
                return folder.getAbsolutePath();
            } else {
                System.err.println("Failed to create folder: " + folder.getAbsolutePath());
            }
        } else {
            System.out.println("Folder already exists: " + folder.getAbsolutePath());
        }

        return null; // Return null in case of failure

    }
    public static String pathBefore(String path) {
        File file = new File(path);

        if (file.exists()) {
            File parent = file.getParentFile();
            if (parent != null) {
                return parent.getAbsolutePath();
            }
        }

        return null; // Return null if the path doesn't exist or there's no parent folder
    }
    public File searchForFile(String directoryPath, String targetFileName) {
        File directory = new File(directoryPath);
        return searchForFile(directory, targetFileName);
    }
    private File searchForFile(File directory, String targetFileName) {
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    File found = searchForFile(file, targetFileName);
                    if (found != null) {
                        return found; // Return the found file if it's in a subdirectory
                    }
                } else {
                    if (file.getName().equals(targetFileName)) {
                        return file; // Return the found file
                    }
                }
            }
        }
        return null; // File not found
    }
    public void saveFileToOutputPath(File foundFile, String outputPath) throws IOException {
        Path sourcePath = foundFile.toPath();
        Path destinationPath = new File(outputPath, foundFile.getName()).toPath();
        Files.copy(sourcePath, destinationPath, StandardCopyOption.REPLACE_EXISTING);
        System.out.println("File saved to: " + destinationPath);
    }

}
