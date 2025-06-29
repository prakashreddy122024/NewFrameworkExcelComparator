package org.example;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelRowColumnCounter {
    public static Map<String, Integer> getSheetRowCount(String filePath) throws Exception {
        Map<String, Integer> sheetRowCount = new HashMap<>();
        try (OPCPackage pkg = OPCPackage.open(filePath)) {
            XSSFReader reader = new XSSFReader(pkg);
           SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    String sheetName = iter.getSheetName();
                    RowCountHandler handler = new RowCountHandler();
                    XMLReader parser = XMLReaderFactory.createXMLReader();
                    parser.setContentHandler(new XSSFSheetXMLHandler(styles, sst, handler, false));
                    parser.parse(new InputSource(stream));
                    sheetRowCount.put(sheetName, handler.getRowCount());
                }
            }
        }
        return sheetRowCount;
    }

    public static Map<String, Integer> getSheetMaxColumnCount(String filePath) throws Exception {
        Map<String, Integer> sheetColCount = new HashMap<>();
        try (OPCPackage pkg = OPCPackage.open(filePath)) {
            XSSFReader reader = new XSSFReader(pkg);
           SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    String sheetName = iter.getSheetName();
                    MaxColumnCountHandler handler = new MaxColumnCountHandler();
                    XMLReader parser = XMLReaderFactory.createXMLReader();
                    parser.setContentHandler(new XSSFSheetXMLHandler(styles, sst, handler, false));
                    parser.parse(new InputSource(stream));
                    sheetColCount.put(sheetName, handler.getMaxColCount());
                }
            }
        }
        return sheetColCount;
    }

    public static void writeHtmlResult(Map<String, Integer> rowOrColCounts, String type) {
        List<String> htmlMessages = new ArrayList<>();
        htmlMessages.add(type + " Counts:");
        if (rowOrColCounts == null || rowOrColCounts.isEmpty()) {
            htmlMessages.add("No " + type.toLowerCase() + " counts found.");
        } else {
            for (Map.Entry<String, Integer> entry : rowOrColCounts.entrySet()) {
                htmlMessages.add("Sheet: " + entry.getKey() + ", " + type + ": " + entry.getValue());
            }
        }
        try {
            ExcelComparisonHtmlWriter.writeComparisonResult(htmlMessages);
        } catch (Exception e) {
            System.err.println("Failed to write HTML result: " + e.getMessage());
        }
    }

    // Handler for counting rows
    private static class RowCountHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private int rowCount = 0;
        @Override
        public void startRow(int rowNum) {
            rowCount++;
        }
        @Override
        public void endRow(int rowNum) {}
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {}
        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {}
        public int getRowCount() { return rowCount; }
    }

    // Handler for finding max column count in a sheet
    private static class MaxColumnCountHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private int maxColCount = 0;
        private int currentColCount = 0;
        @Override
        public void startRow(int rowNum) {
            currentColCount = 0;
        }
        @Override
        public void endRow(int rowNum) {
            if (currentColCount > maxColCount) {
                maxColCount = currentColCount;
            }
        }
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            currentColCount++;
        }
        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {}
        public int getMaxColCount() { return maxColCount; }
    }
}