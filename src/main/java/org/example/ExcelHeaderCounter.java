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

public class ExcelHeaderCounter {
    public static Map<String, List<String>> getSheetHeaders(String filePath) throws Exception {
        Map<String, List<String>> sheetHeaders = new LinkedHashMap<>();
        try (OPCPackage pkg = OPCPackage.open(filePath)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    String sheetName = iter.getSheetName();
                    HeaderHandler handler = new HeaderHandler();
                    XMLReader parser = XMLReaderFactory.createXMLReader();
                    parser.setContentHandler(new XSSFSheetXMLHandler(styles, sst, handler, false));
                    parser.parse(new InputSource(stream));
                    sheetHeaders.put(sheetName, handler.getHeaders());
                }
            }
        }
        return sheetHeaders;
    }

    public static Map<String, Integer> getSheetHeaderCount(String filePath) throws Exception {
        Map<String, Integer> headerCounts = new LinkedHashMap<>();
        Map<String, List<String>> headers = getSheetHeaders(filePath);
        for (Map.Entry<String, List<String>> entry : headers.entrySet()) {
            headerCounts.put(entry.getKey(), entry.getValue().size());
        }
        return headerCounts;
    }

    // Handler to capture the first row (header) of each sheet
    private static class HeaderHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private List<String> headers = new ArrayList<>();
        private boolean isFirstRow = true;
        @Override
        public void startRow(int rowNum) {
            if (rowNum == 0) {
                headers = new ArrayList<>();
                isFirstRow = true;
            } else {
                isFirstRow = false;
            }
        }
        @Override
        public void endRow(int rowNum) {}
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            if (isFirstRow) {
                headers.add(formattedValue);
            }
        }
        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {}
        public List<String> getHeaders() { return headers; }
    }

    public static void printHeaderComparison(String file1, String file2, String shortFile1, String shortFile2) {
        try {
            Map<String, List<String>> file1Headers = getSheetHeaders(file1);
            Map<String, List<String>> file2Headers = getSheetHeaders(file2);
            Set<String> allSheets = new HashSet<>();
            allSheets.addAll(file1Headers.keySet());
            allSheets.addAll(file2Headers.keySet());
            for (String sheet : allSheets) {
                List<String> h1 = file1Headers.get(sheet);
                List<String> h2 = file2Headers.get(sheet);
                int c1 = h1 == null ? 0 : h1.size();
                int c2 = h2 == null ? 0 : h2.size();
                System.out.println("Sheet: " + sheet + " | " + shortFile1 + " header count: " + c1 + ", " + shortFile2 + " header count: " + c2);
                if (h1 != null && h2 != null) {
                    Set<String> onlyIn1 = new HashSet<>(h1);
                    onlyIn1.removeAll(h2);
                    Set<String> onlyIn2 = new HashSet<>(h2);
                    onlyIn2.removeAll(h1);
                    if (!onlyIn1.isEmpty()) {
                        System.out.println("  Headers only in " + shortFile1 + ": " + onlyIn1);
                    }
                    if (!onlyIn2.isEmpty()) {
                        System.out.println("  Headers only in " + shortFile2 + ": " + onlyIn2);
                    }
                    if (onlyIn1.isEmpty() && onlyIn2.isEmpty()) {
                        System.out.println("  Both files have the same headers in sheet: " + sheet);
                    }
                } else if (h1 == null) {
                    System.out.println("  Sheet '" + sheet + "' missing in " + shortFile1);
                } else if (h2 == null) {
                    System.out.println("  Sheet '" + sheet + "' missing in " + shortFile2);
                }
            }
        } catch (Exception e) {
            System.err.println("Error comparing headers: " + e.getMessage());
        }
    }

    public static void writeHtmlResult(Map<String, List<String>> sheetHeaders) {
        List<String> htmlMessages = new ArrayList<>();
        htmlMessages.add("Sheet Headers:");
        if (sheetHeaders == null || sheetHeaders.isEmpty()) {
            htmlMessages.add("No headers found in any sheet.");
        } else {
            for (Map.Entry<String, List<String>> entry : sheetHeaders.entrySet()) {
                htmlMessages.add("Sheet: " + entry.getKey() + ", Headers: " + entry.getValue());
            }
        }
        try {
            ExcelComparisonHtmlWriter.writeComparisonResult(htmlMessages);
        } catch (Exception e) {
            System.err.println("Failed to write HTML result: " + e.getMessage());
        }
    }
}
