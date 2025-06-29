package org.example;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;

public class ExcelFileValidator {
    public static boolean isValidExcelFile(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis)) {
            // Try to open the workbook, if no exception, it's valid
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}

