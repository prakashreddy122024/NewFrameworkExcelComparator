package org.example;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;

import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

public class ExcelSheetValidator {
    public static boolean validateSheetCount(String file1, String file2) {
        try {
            int count1 = getSheetCount(file1);
            int count2 = getSheetCount(file2);
            return count1 == count2;
        } catch (Exception e) {
            System.err.println("Error validating sheet count: " + e.getMessage());
            return false;
        }
    }

    public static int getSheetCount(String filePath) throws Exception {
        try (OPCPackage pkg = OPCPackage.open(filePath)) {
            XSSFReader reader = new XSSFReader(pkg);
            int count = 0;
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (iter.hasNext()) {
                iter.next();
                count++;
            }
            return count;
        }
    }

    public static void printSheetDetails(String file1, String file2) {
        try {
            int count1 = getSheetCount(file1);
            int count2 = getSheetCount(file2);
            System.out.println(file1+": sheet count: " + count1);
            System.out.println(file2+": sheet count: " + count2);

            Set<String> sheets1 = getSheetNames(file1);
            Set<String> sheets2 = getSheetNames(file2);

            Set<String> onlyIn1 = new HashSet<>(sheets1);
            onlyIn1.removeAll(sheets2);
            Set<String> onlyIn2 = new HashSet<>(sheets2);
            onlyIn2.removeAll(sheets1);

            if (!onlyIn1.isEmpty()) {
                System.out.println("Sheets only in "+file1+": " + onlyIn1);
            }
            if (!onlyIn2.isEmpty()) {
                System.out.println("Sheets only in "+file2+": " + onlyIn2);
            }
            if (onlyIn1.isEmpty() && onlyIn2.isEmpty()) {
                System.out.println("Both files have the same sheet names.");
            }
        } catch (Exception e) {
            System.err.println("Error printing sheet details: " + e.getMessage());
        }
    }

    public static Set<String> getSheetNames(String filePath) throws Exception {
        Set<String> names = new LinkedHashSet<>();
        try (OPCPackage pkg = OPCPackage.open(filePath)) {
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (iter.hasNext()) {
                iter.next();
                names.add(iter.getSheetName());
            }
        }
        return names;
    }

    /**
     * Returns sheet details as a list of strings for both console and HTML reporting.
     */
    public static java.util.List<String> getSheetDetails(String file1, String file2) {
        java.util.List<String> details = new java.util.ArrayList<>();
        try {
            int count1 = getSheetCount(file1);
            int count2 = getSheetCount(file2);
            details.add(file1 + ": sheet count: " + count1);
            details.add(file2 + ": sheet count: " + count2);

            Set<String> sheets1 = getSheetNames(file1);
            Set<String> sheets2 = getSheetNames(file2);

            Set<String> onlyIn1 = new HashSet<>(sheets1);
            onlyIn1.removeAll(sheets2);
            Set<String> onlyIn2 = new HashSet<>(sheets2);
            onlyIn2.removeAll(sheets1);

            if (!onlyIn1.isEmpty()) {
                details.add("Sheets only in " + file1 + ": " + onlyIn1);
            }
            if (!onlyIn2.isEmpty()) {
                details.add("Sheets only in " + file2 + ": " + onlyIn2);
            }
            if (onlyIn1.isEmpty() && onlyIn2.isEmpty()) {
                details.add("Both files have the same sheet names.");
            }
        } catch (Exception e) {
            details.add("Error printing sheet details: " + e.getMessage());
        }
        return details;
    }

    public static List<String> getSheetDetails(String shortFile1, String shortFile2, String file1, String file2) {
        List<String> details = new java.util.ArrayList<>();
        try {
            int count1 = getSheetCount(file1);
            int count2 = getSheetCount(file2);
            details.add(shortFile1 + ": sheet count: " + count1);
            details.add(shortFile2 + ": sheet count: " + count2);

            Set<String> sheets1 = getSheetNames(file1);
            Set<String> sheets2 = getSheetNames(file2);

            Set<String> onlyIn1 = new HashSet<>(sheets1);
            onlyIn1.removeAll(sheets2);
            Set<String> onlyIn2 = new HashSet<>(sheets2);
            onlyIn2.removeAll(sheets1);

            if (!onlyIn1.isEmpty()) {
                details.add("Sheets only in " + shortFile1 + ": " + onlyIn1);
            }
            if (!onlyIn2.isEmpty()) {
                details.add("Sheets only in " + shortFile2 + ": " + onlyIn2);
            }
            if (onlyIn1.isEmpty() && onlyIn2.isEmpty()) {
                details.add("Both files have the same sheet names.");
            }
        } catch (Exception e) {
            details.add("Error printing sheet details: " + e.getMessage());
        }
        return details;
    }

    public static void writeHtmlResult(String file1, String file2) {
        List<String> details = getSheetDetails(file1, file2);
        try {
            ExcelComparisonHtmlWriter.writeComparisonResult(details);
        } catch (Exception e) {
            System.err.println("Failed to write HTML result: " + e.getMessage());
        }
    }
}
