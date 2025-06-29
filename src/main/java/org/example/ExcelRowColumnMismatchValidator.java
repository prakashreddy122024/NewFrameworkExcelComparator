package org.example;

import java.util.*;

/**
 * Validates mismatching data for rows and columns when headers are the same in both sheets.
 */
public class ExcelRowColumnMismatchValidator {
    /**
     * Compares two Excel data maps and returns a list of mismatches for rows and columns where headers are the same.
     *
     * @param file1Data Map of row keys to row data for file 1
     * @param file2Data Map of row keys to row data for file 2
     * @param headers List of header names (assumed to be the same for both files)
     * @return List of mismatch descriptions
     */
    public static List<String> validateRowColumnMismatches(Map<String, List<String>> file1Data, Map<String, List<String>> file2Data, List<String> headers) {
        List<String> mismatches = new ArrayList<>();
        Set<String> commonKeys = new HashSet<>(file1Data.keySet());
        commonKeys.retainAll(file2Data.keySet());
        for (String key : commonKeys) {
            List<String> row1 = file1Data.get(key);
            List<String> row2 = file2Data.get(key);
            int maxCols = Math.max(row1.size(), row2.size());
            for (int i = 0; i < maxCols; i++) {
                String val1 = i < row1.size() ? row1.get(i) : "";
                String val2 = i < row2.size() ? row2.get(i) : "";
                String header = (i < headers.size()) ? headers.get(i) : ("Column " + (i+1));
                if (!Objects.equals(val1, val2)) {
                    mismatches.add("Mismatch at Row: '" + key + "', Header: '" + header + "', File1: '" + val1 + "', File2: '" + val2 + "'");
                }
            }
        }
        return mismatches;
    }

    public static void writeHtmlResult(List<String> mismatches) {
        List<String> htmlMessages = new ArrayList<>();
        htmlMessages.add("Row/Column Mismatches (when headers are the same):");
        if (mismatches == null || mismatches.isEmpty()) {
            htmlMessages.add("No row/column mismatches found.");
        } else {
            htmlMessages.addAll(mismatches);
        }
        try {
            ExcelComparisonHtmlWriter.writeComparisonResult(htmlMessages);
        } catch (Exception e) {
            System.err.println("Failed to write HTML result: " + e.getMessage());
        }
    }
}
