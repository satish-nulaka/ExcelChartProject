package com.example.excelchart;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;

public class SimpleExcelTest {
    
    public static void main(String[] args) {
        try {
            File excelFile = new File("src/main/resources/data/Ethos - Dog - Report (3).xlsx");
            
            if (!excelFile.exists()) {
                System.err.println("Excel file not found at: " + excelFile.getAbsolutePath());
                return;
            }
            
            System.out.println("Reading Excel file: " + excelFile.getName());
            
            try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFile))) {
                int sheetCount = workbook.getNumberOfSheets();
                System.out.println("Found " + sheetCount + " worksheets");
                
                for (int i = 0; i < sheetCount; i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();
                    System.out.println("\n=== Sheet " + (i+1) + ": " + sheetName + " ===");
                    
                    // Show first 10 rows
                    for (int rowIndex = 0; rowIndex < Math.min(10, sheet.getLastRowNum() + 1); rowIndex++) {
                        Row row = sheet.getRow(rowIndex);
                        if (row != null) {
                            System.out.print("Row " + rowIndex + ": ");
                            for (int cellIndex = 0; cellIndex < Math.min(10, row.getLastCellNum()); cellIndex++) {
                                Cell cell = row.getCell(cellIndex);
                                if (cell != null) {
                                    String cellValue = getCellValueAsString(cell);
                                    System.out.print("[" + cellValue + "] ");
                                } else {
                                    System.out.print("[null] ");
                                }
                            }
                            System.out.println();
                        }
                    }
                }
            }
            
        } catch (Exception e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return null;
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
} 