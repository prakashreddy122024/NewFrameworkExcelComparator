package org.example;

import java.util.*;

public class ExcelExtraRowsColumnsValidator {
    /**
     * Compares two Excel data maps and returns lists of extra rows and columns in each file.
     * @param file1Data Map of row keys to row data for file 1
     * @param file2Data Map of row keys to row data for file 2
     * @return Map with keys: "extraRowsInFile1", "extraRowsInFile2", "extraColumnsInFile1", "extraColumnsInFile2"
     */
    public static Map<String, List<String>> validateExtraRowsAndColumns(Map<String, List<String>> file1Data, Map<String, List<String>> file2Data) {
        Map<String, List<String>> result = new HashMap<>();
        List<String> extraRowsInFile1 = new ArrayList<>();
        List<String> extraRowsInFile2 = new ArrayList<>();
        List<String> extraColumnsInFile1 = new ArrayList<>();
        List<String> extraColumnsInFile2 = new ArrayList<>();

        // Find extra rows
        for (String key : file1Data.keySet()) {
            if (!file2Data.containsKey(key)) {
                extraRowsInFile1.add(key);
            }
        }
        for (String key : file2Data.keySet()) {
            if (!file1Data.containsKey(key)) {
                extraRowsInFile2.add(key);
            }
        }

        // Find extra columns for matching rows
        Set<String> commonKeys = new HashSet<>(file1Data.keySet());
        commonKeys.retainAll(file2Data.keySet());
        for (String key : commonKeys) {
            List<String> row1 = file1Data.get(key);
            List<String> row2 = file2Data.get(key);
            if (row1.size() > row2.size()) {
                for (int i = row2.size(); i < row1.size(); i++) {
                    extraColumnsInFile1.add("Row '" + key + "', Column: " + (i+1));
                }
            } else if (row2.size() > row1.size()) {
                for (int i = row1.size(); i < row2.size(); i++) {
                    extraColumnsInFile2.add("Row '" + key + "', Column: " + (i+1));
                }
            }
        }

        result.put("extraRowsInFile1", extraRowsInFile1);
        result.put("extraRowsInFile2", extraRowsInFile2);
        result.put("extraColumnsInFile1", extraColumnsInFile1);
        result.put("extraColumnsInFile2", extraColumnsInFile2);
        return result;
    }

    public static void writeHtmlResult(Map<String, List<String>> validationResult) {
        List<String> htmlMessages = new ArrayList<>();
        htmlMessages.add("Extra Rows and Columns Validation Result:");
        htmlMessages.add("Extra Rows in File 1: " + validationResult.getOrDefault("extraRowsInFile1", Collections.emptyList()));
        htmlMessages.add("Extra Rows in File 2: " + validationResult.getOrDefault("extraRowsInFile2", Collections.emptyList()));
        htmlMessages.add("Extra Columns in File 1: " + validationResult.getOrDefault("extraColumnsInFile1", Collections.emptyList()));
        htmlMessages.add("Extra Columns in File 2: " + validationResult.getOrDefault("extraColumnsInFile2", Collections.emptyList()));
        try {
            ExcelComparisonHtmlWriter.writeComparisonResult(htmlMessages);
        } catch (Exception e) {
            System.err.println("Failed to write HTML result: " + e.getMessage());
        }
    }
}
