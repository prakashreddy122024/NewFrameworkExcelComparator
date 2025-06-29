package org.example;

public class ExcelFilePathProvider {
    public static String[] getExcelFilePaths(String[] args) {
        // You can set your default Excel file paths here
        String defaultFile1 = "src/main/resources/BEFORE DATA 5.0.xlsx";
        String defaultFile2 = "src/main/resources/AFTER DATA 5.0.xlsx";
        if (args.length >= 2) {
            return new String[]{args[0], args[1]};
        }
        // Always use default paths if no arguments are provided
        return new String[]{defaultFile1, defaultFile2};
    }
}
