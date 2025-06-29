package org.example;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.*;
import java.io.IOException;

public class ExcelComparator {
    public static void compareExcelFiles(String file1, String file2) throws IOException {
        try {
            // Get short file names for display
            String shortFile1 = getShortName(file1);
            String shortFile2 = getShortName(file2);

            // Print and collect sheet details for both console and HTML
            List<String> sheetDetails = ExcelSheetValidator.getSheetDetails(shortFile1, shortFile2, file1, file2);
            for (String detail : sheetDetails) {
                System.out.println(detail);
            }

            // Validate sheet count before comparing content
            if (!ExcelSheetValidator.validateSheetCount(file1, file2)) {
                ExcelComparisonHtmlWriter.writeComparisonResult(sheetDetails);
                return;
            }

            Map<String, List<String>> file1Data = readExcelToMap(file1);
            Map<String, List<String>> file2Data = readExcelToMap(file2);
            List<String> differences = new ArrayList<>();

            Set<String> allKeys = new HashSet<>();
            allKeys.addAll(file1Data.keySet());
            allKeys.addAll(file2Data.keySet());

            for (String key : allKeys) {
                List<String> row1 = file1Data.get(key);
                List<String> row2 = file2Data.get(key);
                if (row1 == null) {
                    differences.add("Row '" + key + "' missing in " + shortFile1);
                } else if (row2 == null) {
                    differences.add("Row '" + key + "' missing in " + shortFile2);
                } else {
                    int maxCols = Math.max(row1.size(), row2.size());
                    for (int j = 0; j < maxCols; j++) {
                        String val1 = j < row1.size() ? row1.get(j) : "";
                        String val2 = j < row2.size() ? row2.get(j) : "";
                        if (!val1.equals(val2)) {
                            differences.add("Difference at Row: '" + key + "', Column: " + (j+1) + ", " + shortFile1 + ": '" + val1 + "', " + shortFile2 + ": '" + val2 + "'");
                        }
                    }
                }
            }
            // If no differences, print and add a message for HTML
            if (differences.isEmpty()) {
                String msg = "Both Excel sheets are identical. No differences found.";
                System.out.println(msg);
                differences.add(msg);
            }
            // Combine sheet details and differences for HTML output
            List<String> htmlOutput = new ArrayList<>(sheetDetails);
            if (!differences.isEmpty() && (differences.size() != 1 || !differences.get(0).equals(htmlOutput.get(htmlOutput.size()-1)))) {
                htmlOutput.addAll(differences);
            }
            ExcelComparisonHtmlWriter.writeComparisonResult(htmlOutput);
        } catch (Exception e) {
            throw new IOException("Error comparing Excel files: " + e.getMessage(), e);
        }
    }

    static String getShortName(String filePath) {
        String name = filePath.replaceAll("\\\\", "/");
        name = name.substring(name.lastIndexOf('/') + 1);
        if (name.toLowerCase().endsWith(".xlsx")) {
            name = name.substring(0, name.length() - 5);
        }
        return name;
    }

    static Map<String, List<String>> readExcelToMap(String filePath) throws Exception {
        Map<String, List<String>> data = new LinkedHashMap<>();
        try (OPCPackage pkg = OPCPackage.open(filePath)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
            int sheetIndex = 0;
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    String sheetName = iter.getSheetName();
                    XMLReader parser = XMLReaderFactory.createXMLReader();
                    parser.setContentHandler(new XSSFSheetXMLHandler(
                        styles, sst, new RowHandler(data, sheetName), false));
                    parser.parse(new InputSource(stream));
                }
                sheetIndex++;
            }
        }
        return data;
    }

    private static class RowHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private final Map<String, List<String>> data;
        private final String sheetName;
        private List<String> currentRow;
        private int currentRowNum;

        public RowHandler(Map<String, List<String>> data, String sheetName) {
            this.data = data;
            this.sheetName = sheetName;
        }

        @Override
        public void startRow(int rowNum) {
            currentRow = new ArrayList<>();
            currentRowNum = rowNum;
        }

        @Override
        public void endRow(int rowNum) {
            String key = sheetName + ":" + (currentRowNum + 1);
            data.put(key, new ArrayList<>(currentRow));
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            currentRow.add(formattedValue);
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
            // Not used
        }
    }

}
