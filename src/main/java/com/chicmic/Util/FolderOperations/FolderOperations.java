package com.chicmic.Util.FolderOperations;

import com.chicmic.engine.MainRunner;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

public class FolderOperations {
    public String createFolder(String folderName, String path) {
        File folder = new File(path, folderName);

        if (folder.exists()) {
            if(MainRunner.autoDeleteFolder) {
                // Delete the folder if it already exists
                boolean deleted = deleteFolder(folder);
                if (!deleted) {
                    System.err.println(getClass().getName() + " : Failed to delete existing folder: " + folder.getAbsolutePath());
                    return null;
                } else {
                    System.out.println("\u001B[31m Folder Deleted: " + folder.getAbsolutePath() + "\u001B[0m");
                }
            }else {
                System.err.println(getClass().getName() + " : Folder already exists: " + folder.getAbsolutePath());
                return null;
            }
        }

        boolean created = folder.mkdirs();
        if (created) {
            System.out.println("Folder created: " + folder.getAbsolutePath());
            return folder.getAbsolutePath();
        } else {
            System.err.println(getClass().getName() + " : Failed to create folder: " + folder.getAbsolutePath());
            return null; // Return null in case of failure
        }
    }

    public boolean deleteFolder(File folder) {
        if (folder.isDirectory()) {
            File[] files = folder.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.isDirectory()) {
                        deleteFolder(file);
                    } else {
                        boolean deleted = file.delete();
                        if (!deleted) {
                            return false;
                        }
                    }
                }
            }
        }
        return folder.delete();
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
                        return found;
                    }
                } else {
                    if (file.getName().equals(targetFileName)) {
                        return file;
                    }
                }
            }
        }
        return null; // File not found
    }
    public void saveFileToOutputPath(File foundFile, String outputPath) throws IOException {
        if(foundFile == null) {
            System.err.println("File not found.");
            return;
        }
        Path sourcePath = foundFile.toPath();
        Path destinationPath = new File(outputPath, foundFile.getName()).toPath();
        Files.copy(sourcePath, destinationPath, StandardCopyOption.REPLACE_EXISTING);
//        System.out.println("File saved to: " + destinationPath);
    }
    public  boolean deleteFile(String filePath) {
        File fileToDelete = new File(filePath);
        if (fileToDelete.exists()) {
            return fileToDelete.delete();
        } else {
            System.err.println("File does not exist.");
            return false;  // File does not exist
        }
    }
}
