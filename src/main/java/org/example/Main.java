package org.example;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.util.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.util.List;
import java.io.InputStream;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws Exception {
        String[] filePaths = ExcelFilePathProvider.getExcelFilePaths(args);
        System.out.println("Excel comparison execution started...");
        IOUtils.setByteArrayMaxOverride(500_000_000); // Set max byte array size to 500MB
        ZipSecureFile.setMinInflateRatio(0.0); // Allow highly compressed files (disable zip bomb detection)
        if (!isValidLargeExcelFile(filePaths[0])) {
            System.err.println("The first Excel file is invalid or corrupted: " + filePaths[0]);
            return;
        }
        if (!isValidLargeExcelFile(filePaths[1])) {
            System.err.println("The second Excel file is invalid or corrupted: " + filePaths[1]);
            return;
        }
        String shortFile1 = ExcelComparator.getShortName(filePaths[0]);
        String shortFile2 = ExcelComparator.getShortName(filePaths[1]);
        try {
            ExcelComparator.compareExcelFiles(filePaths[0], filePaths[1]);
        } catch (Exception e) {
            System.err.println("Error comparing Excel files: " + e.getMessage());
        }

        // Print row and column counts for each sheet in both files
        try {
            System.out.println("\nRow and Column Counts for " + shortFile1 + ":");
            Map<String, Integer> file1Rows = ExcelRowColumnCounter.getSheetRowCount(filePaths[0]);
            Map<String, Integer> file1Cols = ExcelRowColumnCounter.getSheetMaxColumnCount(filePaths[0]);
            for (String sheet : file1Rows.keySet()) {
                System.out.println("Sheet: " + sheet + ", Rows: " + file1Rows.get(sheet) + ", Columns: " + file1Cols.get(sheet));
            }
            System.out.println("\nRow and Column Counts for " + shortFile2 + ":");
            Map<String, Integer> file2Rows = ExcelRowColumnCounter.getSheetRowCount(filePaths[1]);
            Map<String, Integer> file2Cols = ExcelRowColumnCounter.getSheetMaxColumnCount(filePaths[1]);
            for (String sheet : file2Rows.keySet()) {
                System.out.println("Sheet: " + sheet + ", Rows: " + file2Rows.get(sheet) + ", Columns: " + file2Cols.get(sheet));
            }
        } catch (Exception e) {
            System.err.println("Error printing row/column counts: " + e.getMessage());
        }

       // --- Row and Column Counts HTML Output ---
        try {
            Map<String, Integer> file1Rows = ExcelRowColumnCounter.getSheetRowCount(filePaths[0]);
            Map<String, Integer> file1Cols = ExcelRowColumnCounter.getSheetMaxColumnCount(filePaths[0]);
            Map<String, Integer> file2Rows = ExcelRowColumnCounter.getSheetRowCount(filePaths[1]);
            Map<String, Integer> file2Cols = ExcelRowColumnCounter.getSheetMaxColumnCount(filePaths[1]);
            ExcelRowColumnCounter.writeHtmlResult(file1Rows, "Row");
            ExcelRowColumnCounter.writeHtmlResult(file1Cols, "Column");
            ExcelRowColumnCounter.writeHtmlResult(file2Rows, "Row");
            ExcelRowColumnCounter.writeHtmlResult(file2Cols, "Column");
        } catch (Exception e) {
            System.err.println("Error writing row/column HTML result: " + e.getMessage());
        }

        // Print header count for each sheet in both files
        try {
            System.out.println("\nHeader Count for " + shortFile1 + ":");
            Map<String, Integer> file1HeaderCount = ExcelHeaderCounter.getSheetHeaderCount(filePaths[0]);
            for (Map.Entry<String, Integer> entry : file1HeaderCount.entrySet()) {
                System.out.println("Sheet: " + entry.getKey() + ", Header Count: " + entry.getValue());
            }
            System.out.println("\nHeader Count for " + shortFile2 + ":");
            Map<String, Integer> file2HeaderCount = ExcelHeaderCounter.getSheetHeaderCount(filePaths[1]);
            for (Map.Entry<String, Integer> entry : file2HeaderCount.entrySet()) {
                System.out.println("Sheet: " + entry.getKey() + ", Header Count: " + entry.getValue());
            }
        } catch (Exception e) {
            System.err.println("Error printing header counts: " + e.getMessage());
        }

        // Print header comparison for each sheet in both files
        try {
            System.out.println("\nHeader Comparison for " + shortFile1 + " and " + shortFile2 + ":");
            ExcelHeaderCounter.printHeaderComparison(filePaths[0], filePaths[1], shortFile1, shortFile2);
        } catch (Exception e) {
            System.err.println("Error printing header comparison: " + e.getMessage());
        }

        // --- Extra Rows and Columns Validation ---
        Map<String, java.util.List<String>> file1Data = ExcelComparator.readExcelToMap(filePaths[0]);
        Map<String, java.util.List<String>> file2Data = ExcelComparator.readExcelToMap(filePaths[1]);
        Map<String, java.util.List<String>> extraValidation = ExcelExtraRowsColumnsValidator.validateExtraRowsAndColumns(file1Data, file2Data);
        ExcelExtraRowsColumnsValidator.writeHtmlResult(extraValidation);
        System.out.println("\nExtra Rows in " + shortFile1 + ": " + extraValidation.get("extraRowsInFile1"));
        System.out.println("Extra Rows in " + shortFile2 + ": " + extraValidation.get("extraRowsInFile2"));
        System.out.println("Extra Columns in " + shortFile1 + ": " + extraValidation.get("extraColumnsInFile1"));
        System.out.println("Extra Columns in " + shortFile2 + ": " + extraValidation.get("extraColumnsInFile2"));

        // --- Row/Column Mismatch Validation when headers are the same ---
        // Assume first row of the first sheet is the header row for both files
        List<String> headers = null;
        if (!file1Data.isEmpty()) {
            String firstKey = file1Data.keySet().iterator().next();
            headers = file1Data.get(firstKey);
        }
        if (headers != null && !headers.isEmpty()) {
            List<String> mismatches = ExcelRowColumnMismatchValidator.validateRowColumnMismatches(file1Data, file2Data, headers);
            ExcelRowColumnMismatchValidator.writeHtmlResult(mismatches);
            System.out.println("\nRow/Column Mismatches (when headers are the same):");
            for (String mismatch : mismatches) {
                System.out.println(mismatch);
            }
        } else {
            System.out.println("\nNo headers found to validate row/column mismatches.");
        }


        System.out.println("Excel comparison execution completed.");
    }

    // Efficient validation for large .xlsx files using POI event API
    public static boolean isValidLargeExcelFile(String filePath) {
        try (OPCPackage pkg = OPCPackage.open(filePath)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    XMLReader parser = XMLReaderFactory.createXMLReader();
                    parser.setContentHandler(new XSSFSheetXMLHandler(
                            styles, sst, new XSSFSheetXMLHandler.SheetContentsHandler() {
                        @Override
                        public void startRow(int rowNum) {
                        }

                        @Override
                        public void endRow(int rowNum) {
                        }

                        @Override
                        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                        }

                        @Override
                        public void headerFooter(String text, boolean isHeader, String tagName) {
                        }
                    }, false));
                    parser.parse(new InputSource(stream));
                }
            }
            return true; // No exception, file is valid
        } catch (Exception e) {
            System.err.println("Validation error for file: " + filePath + " - " + e.getMessage());
            return false;
        }
    }
}
