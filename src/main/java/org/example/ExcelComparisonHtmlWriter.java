package org.example;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

public class ExcelComparisonHtmlWriter {
    private static final String htmlPath = "src/main/resources/comparison_result.html";

    private static String getCssClass(String line) {
        String lower = line.toLowerCase();
        if (lower.contains("identical") || lower.contains("completed") || lower.contains("same sheet names") || lower.contains("same headers")) {
            return "success";
        } else if (lower.contains("difference") || lower.contains("missing") || lower.contains("only in") || lower.contains("error")) {
            return "diff";
        } else if (lower.contains("row and column counts") || lower.contains("header count") || lower.contains("header comparison") || lower.contains("execution started") || lower.contains("execution completed")) {
            return "section";
        }
        return "info";
    }

    public static void writeComparisonResult(List<String> differences) throws IOException {
        String currentDateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(htmlPath, false))) {
            writer.write("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n    <meta charset=\"UTF-8\">\n    <title>Excel Comparison Result</title>\n    <meta name='viewport' content='width=device-width, initial-scale=1'>\n    <link href='https://fonts.googleapis.com/css?family=Roboto:400,700&display=swap' rel='stylesheet'>\n    <script src='https://cdn.jsdelivr.net/npm/chart.js'></script>\n    <style>\n        body { font-family: 'Roboto', Arial, sans-serif; margin: 0; background: #f4f6f8; color: #222; }\n        .container { max-width: 900px; margin: 30px auto; background: #fff; border-radius: 10px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); padding: 32px 28px 28px 28px; }\n        h1 { text-align: center; color: #2a5298; margin-bottom: 10px; }\n        .datetime { text-align: center; font-size: 1.15em; color: #2a5298; margin: 18px 0 18px 0; font-weight: bold; }\n        .summary { text-align: center; margin-bottom: 24px; font-size: 1.1em; color: #555; }\n        .legend { display: flex; justify-content: center; gap: 18px; margin-bottom: 18px; }\n        .legend span { display: flex; align-items: center; gap: 6px; font-size: 0.98em; }\n        .legend .dot { width: 16px; height: 16px; border-radius: 50%; display: inline-block; }\n        .info { color: #00529B; background-color: #BDE5F8; }\n        .success { color: #4F8A10; background-color: #DFF2BF; }\n        .diff { color: #D8000C; background-color: #FFBABA; }\n        .section { color: #222; background-color: #E0E0E0; font-weight: bold; }\n        .pre { white-space: pre-wrap; font-family: inherit; padding: 8px 12px; border-radius: 5px; margin-bottom: 7px; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }\n        .msg-table { width: 100%; border-collapse: collapse; margin-top: 18px; }\n        .msg-table th, .msg-table td { padding: 8px 10px; border-bottom: 1px solid #e0e0e0; text-align: left; }\n        .msg-table th { background: #f0f4fa; color: #2a5298; }\n        .msg-table tr:last-child td { border-bottom: none; }\n        .chart-section { margin: 30px 0 18px 0; text-align: center; }\n        @media (max-width: 600px) { .container { padding: 10px; } .msg-table th, .msg-table td { padding: 6px 4px; font-size: 0.98em; } }\n    </style>\n</head>\n<body>\n    <div class='container'>\n        <h1>Excel Comparison Result</h1>\n        <div class='datetime'>" + currentDateTime + "</div>\n        <div class='summary'>Below is a detailed and visual summary of your Excel file comparison. Differences, extra rows/columns, and mismatches are highlighted for easy review.</div>\n        <div class='legend'>\n            <span><span class='dot info'></span>Info</span>\n            <span><span class='dot success'></span>Success</span>\n            <span><span class='dot diff'></span>Difference</span>\n            <span><span class='dot section'></span>Section</span>\n        </div>\n        <div class='chart-section'>\n            <canvas id='diffChart' width='700' height='280'></canvas>\n        </div>\n        <table class='msg-table'>\n            <thead><tr><th>#</th><th>Message</th></tr></thead>\n            <tbody>\n");

            int infoCount = 0, successCount = 0, diffCount = 0, sectionCount = 0;
            int idx = 1;
            if (differences == null || differences.isEmpty()) {
                writer.write("<tr><td>1</td><td class='success pre'>No differences or validations to display.</td></tr>\n");
                successCount = 1;
            } else {
                for (String line : differences) {
                    String cssClass = getCssClass(line);
                    switch (cssClass) {
                        case "success": successCount++; break;
                        case "diff": diffCount++; break;
                        case "section": sectionCount++; break;
                        default: infoCount++;
                    }
                    writer.write("<tr><td>" + idx++ + "</td><td class='" + cssClass + " pre'>" + escapeHtml(line) + "</td></tr>\n");
                }
            }
            writer.write("            </tbody>\n        </table>\n    </div>\n    <script>\n    const ctx = document.getElementById('diffChart').getContext('2d');\n    const diffChart = new Chart(ctx, {\n        type: 'bar',\n        data: {\n            labels: ['Info', 'Success', 'Diff', 'Section'],\n            datasets: [{\n                label: 'Message Count',\n                data: [" + infoCount + ", " + successCount + ", " + diffCount + ", " + sectionCount + "],\n                backgroundColor: [\n                    '#BDE5F8',\n                    '#DFF2BF',\n                    '#FFBABA',\n                    '#E0E0E0'\n                ],\n                borderColor: [\n                    '#00529B',\n                    '#4F8A10',\n                    '#D8000C',\n                    '#222'\n                ],\n                borderWidth: 2\n            }]\n        },\n        options: {\n            responsive: true,\n            plugins: {\n                legend: { display: false },\n                title: { display: true, text: 'Comparison Message Type Distribution', color: '#2a5298', font: { size: 18 } }\n            },\n            scales: {\n                y: { beginAtZero: true, ticks: { color: '#2a5298', font: { size: 14 } } },\n                x: { ticks: { color: '#2a5298', font: { size: 14 } } }\n            }\n        }\n    });\n    </script>\n</body>\n</html>");
        }
    }

    /**
     * Collects and writes all validation results to the HTML report using ExcelComparisonHtmlWriter.
     * Call this after all validations to ensure the HTML report is comprehensive.
     */
    public static void writeAllValidationsToHtml(List<String> sheetDetails, List<String> differences, Map<String, List<String>> extraValidation, List<String> mismatches, Map<String, Integer> rowCounts, Map<String, Integer> colCounts, Map<String, List<String>> headers) {
        List<String> htmlMessages = new ArrayList<>();
        if (sheetDetails != null && !sheetDetails.isEmpty()) {
            htmlMessages.add("Sheet Details:");
            htmlMessages.addAll(sheetDetails);
        }
        if (rowCounts != null && !rowCounts.isEmpty()) {
            htmlMessages.add("Row Counts:");
            for (Map.Entry<String, Integer> entry : rowCounts.entrySet()) {
                htmlMessages.add("Sheet: " + entry.getKey() + ", Rows: " + entry.getValue());
            }
        }
        if (colCounts != null && !colCounts.isEmpty()) {
            htmlMessages.add("Column Counts:");
            for (Map.Entry<String, Integer> entry : colCounts.entrySet()) {
                htmlMessages.add("Sheet: " + entry.getKey() + ", Columns: " + entry.getValue());
            }
        }
        if (headers != null && !headers.isEmpty()) {
            htmlMessages.add("Headers:");
            for (Map.Entry<String, List<String>> entry : headers.entrySet()) {
                htmlMessages.add("Sheet: " + entry.getKey() + ", Headers: " + entry.getValue());
            }
        }
        if (extraValidation != null && !extraValidation.isEmpty()) {
            htmlMessages.add("Extra Rows and Columns Validation Result:");
            htmlMessages.add("Extra Rows in File 1: " + extraValidation.getOrDefault("extraRowsInFile1", Collections.emptyList()));
            htmlMessages.add("Extra Rows in File 2: " + extraValidation.getOrDefault("extraRowsInFile2", Collections.emptyList()));
            htmlMessages.add("Extra Columns in File 1: " + extraValidation.getOrDefault("extraColumnsInFile1", Collections.emptyList()));
            htmlMessages.add("Extra Columns in File 2: " + extraValidation.getOrDefault("extraColumnsInFile2", Collections.emptyList()));
        }
        if (differences != null && !differences.isEmpty()) {
            htmlMessages.add("Differences:");
            htmlMessages.addAll(differences);
        }
        if (mismatches != null && !mismatches.isEmpty()) {
            htmlMessages.add("Row/Column Mismatches (when headers are the same):");
            htmlMessages.addAll(mismatches);
        }
        try {
            ExcelComparisonHtmlWriter.writeComparisonResult(htmlMessages);
        } catch (Exception e) {
            System.err.println("Failed to write all validations to HTML: " + e.getMessage());
        }
    }

    private static String escapeHtml(String s) {
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&#39;");
    }
}
